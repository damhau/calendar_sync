"""CLI entry point for Calendar Sync application."""

import argparse
import sys
from datetime import datetime, timedelta
from pathlib import Path

# Initialize SSL truststore early, before any HTTPS imports
from .utils.ssl_utils import init_ssl
init_ssl()

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
            browser=account.browser,
            use_browser_api=account.use_browser_api,
        )
        # Build a minimal EWSConfig-like object from account
        from .config import EWSConfig

        # Use model_construct to bypass environment variable loading
        ews_cfg = EWSConfig.model_construct(
            server_url=account.server_url,
            primary_email=account.primary_email,
            auth_method="selenium",
            cookie_file=account.cookie_file,
        )
        return EWSSeleniumReader(selenium_auth, ews_cfg)
    elif account.type in ("m365", "m365_read"):
        # Use model_construct to bypass environment variable loading
        m365_cfg = M365Config.model_construct(
            tenant_id=account.tenant_id,
            client_id=account.client_id,
            client_secret=account.client_secret,
            primary_email=account.primary_email,
        )
        m365_auth = M365AuthProvider(m365_cfg, cache_manager)
        return M365CalendarReader(m365_auth, primary_email=account.primary_email)
    else:
        raise ValueError(f"Unknown account type: {account.type}")


def _filter_events_by_day(events: list, account, logger=None) -> list:
    """Filter events based on include_days and exclude_days settings."""
    import logging
    log = logger or logging.getLogger(__name__)

    if not account.include_days and not account.exclude_days:
        return events

    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    filtered = []
    excluded_count = 0

    for event in events:
        # Get the day of week (0=Monday, 6=Sunday)
        day_of_week = event.start.weekday()
        day_name = day_names[day_of_week]

        # If include_days is set, only include those days
        if account.include_days:
            if day_of_week not in account.include_days:
                log.info(f"  ⏭️  Skipping '{event.subject}' - {day_name} not in include_days")
                excluded_count += 1
                continue

        # If exclude_days is set, exclude those days
        if account.exclude_days:
            if day_of_week in account.exclude_days:
                log.info(f"  ⏭️  Skipping '{event.subject}' - {day_name} in exclude_days")
                excluded_count += 1
                continue

        filtered.append(event)

    if excluded_count > 0:
        log.info(f"  📅 Day filter: {excluded_count} event(s) excluded, {len(filtered)} remaining")

    return filtered


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
        "--date",
        type=str,
        default=None,
        help="Sync only a specific date (YYYY-MM-DD format, e.g., 2026-02-04)",
    )
    parser.add_argument(
        "--start-date",
        type=str,
        default=None,
        help="Start date for sync range (YYYY-MM-DD format)",
    )
    parser.add_argument(
        "--end-date",
        type=str,
        default=None,
        help="End date for sync range (YYYY-MM-DD format)",
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
        if args.date:
            # Single day mode
            try:
                import pytz
                day = datetime.strptime(args.date, "%Y-%m-%d")
                start = day.replace(hour=0, minute=0, second=0, tzinfo=pytz.utc)
                end = day.replace(hour=23, minute=59, second=59, tzinfo=pytz.utc)
                logger.info(f"Syncing single day: {args.date}")
            except ValueError:
                logger.error(f"Invalid date format: {args.date}. Use YYYY-MM-DD (e.g., 2026-02-04)")
                return 1
        elif args.start_date or args.end_date:
            # Custom date range mode
            import pytz
            try:
                if args.start_date:
                    start = datetime.strptime(args.start_date, "%Y-%m-%d")
                    start = start.replace(hour=0, minute=0, second=0, tzinfo=pytz.utc)
                else:
                    start = datetime.now(pytz.utc).replace(hour=0, minute=0, second=0, microsecond=0)

                if args.end_date:
                    end = datetime.strptime(args.end_date, "%Y-%m-%d")
                    end = end.replace(hour=23, minute=59, second=59, tzinfo=pytz.utc)
                else:
                    end = start + timedelta(days=7)

                logger.info(f"Syncing date range: {start.date()} to {end.date()}")
            except ValueError as e:
                logger.error(f"Invalid date format. Use YYYY-MM-DD (e.g., 2026-02-04). Error: {e}")
                return 1
        else:
            # Default: use lookback/lookahead
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

                # Apply day filtering (include_days / exclude_days)
                events = _filter_events_by_day(events, account)

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

                # Apply day filtering (include_days / exclude_days)
                events = _filter_events_by_day(events, account)

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

                # Use model_construct to bypass environment variable loading
                m365_cfg = M365Config.model_construct(
                    tenant_id=target_account.tenant_id,
                    client_id=target_account.client_id,
                    client_secret=target_account.client_secret,
                    primary_email=target_account.primary_email,
                )
                m365_auth = M365AuthProvider(m365_cfg, cache_manager)
                target_writer = M365CalendarWriter(m365_auth, primary_email=target_account.primary_email)
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
