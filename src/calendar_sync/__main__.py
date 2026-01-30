"""CLI entry point for Calendar Sync application."""

import argparse
import sys
from pathlib import Path

from .auth.msal_auth import M365AuthProvider
from .auth.selenium_auth import SeleniumEWSAuth
from .auth.token_cache import TokenCacheManager
from .config import M365Config, config, sync_config
from .readers.ews_selenium_reader import EWSSeleniumReader
from .readers.m365_reader import M365CalendarReader
from .utils.date_utils import get_sync_window
from .utils.exceptions import CalendarSyncError
from .utils.logging import setup_logging


def _create_reader(account, cache_manager):
    """Create a calendar reader from an AccountConfig."""
    if account.type == "ews_selenium":
        if not account.server_url:
            raise ValueError(f"Account '{account.name}' requires server_url")
        base_url = account.server_url.split("/EWS")[0]
        selenium_auth = SeleniumEWSAuth(
            base_url=base_url,
            cookie_file=account.cookie_file,
            required_cookies=account.required_cookies,
        )
        # Build a minimal EWSConfig-like object from account
        from .config import EWSConfig

        ews_cfg = EWSConfig(
            server_url=account.server_url,
            primary_email=account.primary_email,
            auth_method="selenium",
            cookie_file=account.cookie_file,
        )
        return EWSSeleniumReader(selenium_auth, ews_cfg)
    elif account.type in ("m365", "m365_read"):
        m365_cfg = M365Config(
            tenant_id=account.tenant_id,
            client_id=account.client_id,
            client_secret=account.client_secret,
        )
        m365_auth = M365AuthProvider(m365_cfg, cache_manager)
        return M365CalendarReader(m365_auth)
    else:
        raise ValueError(f"Unknown account type: {account.type}")


def main() -> int:
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Calendar Sync - Synchronize calendars between Exchange EWS and M365"
    )
    parser.add_argument(
        "--source",
        nargs="*",
        help="Source account name(s) from sync_config.yaml (default: all configured sources)",
    )
    parser.add_argument(
        "--target",
        type=str,
        help="Target account name (default: from sync_config.yaml)",
    )
    parser.add_argument(
        "--list-calendars",
        action="store_true",
        help="List available calendars",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Preview events without syncing",
    )
    parser.add_argument(
        "--sync",
        action="store_true",
        help="Perform sync operation",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Dry run (show what would be synced)",
    )
    parser.add_argument(
        "--clear-cache",
        action="store_true",
        help="Clear authentication token cache",
    )
    parser.add_argument(
        "--lookback",
        type=int,
        default=None,
        help="Days to look back (overrides config)",
    )
    parser.add_argument(
        "--lookahead",
        type=int,
        default=None,
        help="Days to look ahead (overrides config)",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Verbose output",
    )

    args = parser.parse_args()

    # Setup logging
    log_level = "DEBUG" if args.verbose else config.log_level
    logger = setup_logging(level=log_level, log_file=config.log_file)

    try:
        # Initialize token cache
        cache_manager = TokenCacheManager(
            cache_location=Path(config.token_cache_path),
            encrypted=config.token_cache_encrypted,
        )

        # Handle clear cache
        if args.clear_cache:
            cache_manager.clear_cache()
            logger.info("Token cache cleared")
            return 0

        # Resolve sources and target from config
        if not sync_config.has_config:
            logger.error("No sync_config.yaml found or no accounts configured")
            return 1

        source_names = args.source if args.source else sync_config.sources
        target_name = args.target or sync_config.target

        if not source_names:
            logger.error("No source accounts specified")
            return 1

        # Validate account names
        for name in source_names:
            if name not in sync_config.accounts:
                logger.error(f"Unknown account: {name}")
                return 1

        # Resolve sync window
        lookback = args.lookback if args.lookback is not None else sync_config.lookback_days
        lookahead = args.lookahead if args.lookahead is not None else sync_config.lookahead_days
        start, end = get_sync_window(lookback, lookahead)

        # List calendars
        if args.list_calendars:
            for name in source_names:
                account = sync_config.accounts[name]
                print(f"\n=== {name} ({account.type}) ===")
                reader = _create_reader(account, cache_manager)
                calendars = reader.list_calendars()
                print(f"Found {len(calendars)} calendar(s):")
                for cal in calendars:
                    print(f"  - {cal.name} (ID: {cal.id})")
                    if cal.owner_email:
                        print(f"    Owner: {cal.owner_email}")
            return 0

        # Preview events
        if args.preview:
            all_events = []
            for name in source_names:
                account = sync_config.accounts[name]
                print(f"\n=== {name} ({account.type}) ===")
                reader = _create_reader(account, cache_manager)
                events = reader.read_events(start_date=start, end_date=end)

                # Apply skip_subjects filter
                events = [
                    e for e in events
                    if e.subject.lower().strip() not in sync_config.skip_subjects
                ]

                # Apply prefix and category
                for e in events:
                    if account.prefix and not e.subject.startswith(account.prefix):
                        e.subject = f"{account.prefix} {e.subject}"
                    if account.category and account.category not in e.categories:
                        e.categories.append(account.category)

                print(f"Found {len(events)} event(s):")
                for event in events:
                    print(f"  - {event.subject}")
                    print(f"    When: {event.start} to {event.end}")
                    print(
                        f"    Location: {event.location.display_name if event.location else 'None'}"
                    )
                    print()
                all_events.extend(events)

            print(f"\nTotal: {len(all_events)} event(s) from {len(source_names)} source(s)")
            return 0

        # Perform sync
        if args.sync:
            if not target_name:
                logger.error("No target account specified (set in sync_config.yaml or --target)")
                return 1
            if target_name not in sync_config.accounts:
                logger.error(f"Unknown target account: {target_name}")
                return 1

            # Read from all sources
            all_events = []
            for name in source_names:
                account = sync_config.accounts[name]
                logger.info(f"Reading from {name} ({account.type})...")
                reader = _create_reader(account, cache_manager)
                events = reader.read_events(start_date=start, end_date=end)

                # Apply skip_subjects filter
                events = [
                    e for e in events
                    if e.subject.lower().strip() not in sync_config.skip_subjects
                ]

                # Apply prefix and category
                for e in events:
                    if account.prefix and not e.subject.startswith(account.prefix):
                        e.subject = f"{account.prefix} {e.subject}"
                    if account.category and account.category not in e.categories:
                        e.categories.append(account.category)

                all_events.extend(events)
                logger.info(f"  {name}: {len(events)} events")

            logger.info(f"Total events to sync: {len(all_events)}")

            # Create target writer
            target_account = sync_config.accounts[target_name]
            if target_account.type.startswith("m365"):
                from .writers.m365_writer import M365CalendarWriter

                m365_cfg = M365Config(
                    tenant_id=target_account.tenant_id,
                    client_id=target_account.client_id,
                    client_secret=target_account.client_secret,
                )
                m365_auth = M365AuthProvider(m365_cfg, cache_manager)
                target_writer = M365CalendarWriter(m365_auth)
            else:
                logger.error(f"Writing to {target_account.type} not supported")
                return 1

            # Ensure categories exist with correct colors
            for name in source_names:
                account = sync_config.accounts[name]
                if account.category:
                    target_writer.ensure_category(account.category, account.color)

            # Fetch existing events for dedup
            existing = target_writer.get_existing_events(start, end)

            if args.dry_run:
                skipped = 0
                would_create = []
                for event in all_events:
                    key = (event.subject, event.start.strftime("%Y-%m-%dT%H:%M"))
                    if key in existing:
                        skipped += 1
                    else:
                        would_create.append(event)
                print(f"\nDry run - {target_name}:")
                print(f"  Would create: {len(would_create)}")
                print(f"  Already exist (skip): {skipped}")
                for event in would_create:
                    print(f"  + {event.subject}")
                    print(f"    When: {event.start} to {event.end}")
                return 0

            # Write events, skipping duplicates
            created = 0
            skipped = 0
            errors = []
            for event in all_events:
                key = (event.subject, event.start.strftime("%Y-%m-%dT%H:%M"))
                if key in existing:
                    skipped += 1
                    continue
                try:
                    target_writer.create_event(event)
                    created += 1
                except Exception as e:
                    error_msg = f"Failed to sync '{event.subject}': {e}"
                    logger.error(error_msg)
                    errors.append(error_msg)

            print(f"\nSync Results:")
            print(f"  Events read: {len(all_events)}")
            print(f"  Events created: {created}")
            print(f"  Events skipped (already exist): {skipped}")
            if errors:
                print(f"\nErrors ({len(errors)}):")
                for err in errors:
                    print(f"  - {err}")
                return 1
            return 0

        # No action specified
        parser.print_help()
        return 0

    except CalendarSyncError as e:
        logger.error(f"Calendar sync error: {e}")
        return 1
    except Exception as e:
        logger.exception(f"Unexpected error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
