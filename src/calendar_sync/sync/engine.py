"""Main calendar synchronization engine."""

import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

from ..readers.base import CalendarReader
from ..utils.date_utils import get_sync_window
from ..writers.base import CalendarWriter
from .strategies import OneWaySyncStrategy, SyncStrategy

logger = logging.getLogger(__name__)


@dataclass
class SyncResult:
    """Result of a sync operation."""

    events_read: int = 0
    events_created: int = 0
    events_updated: int = 0
    events_skipped: int = 0
    errors: list[str] = field(default_factory=list)


class SyncEngine:
    """Main calendar synchronization engine."""

    def __init__(
        self,
        source_reader: CalendarReader,
        target_writer: Optional[CalendarWriter] = None,
        strategy: Optional[SyncStrategy] = None,
    ):
        """
        Initialize sync engine.

        Args:
            source_reader: Calendar reader for source system
            target_writer: Calendar writer for target system (optional)
            strategy: Sync strategy (defaults to OneWaySyncStrategy)
        """
        self.source_reader = source_reader
        self.target_writer = target_writer
        self.strategy = strategy or OneWaySyncStrategy()

    def sync(
        self,
        source_calendar_id: Optional[str] = None,
        target_calendar_id: Optional[str] = None,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
        dry_run: bool = False,
    ) -> SyncResult:
        """
        Synchronize events from source to target.

        Args:
            source_calendar_id: Source calendar ID (None for default)
            target_calendar_id: Target calendar ID (None for default)
            start_date: Start of sync window
            end_date: End of sync window
            dry_run: If True, only report what would be synced

        Returns:
            SyncResult with statistics
        """
        result = SyncResult()

        try:
            # Get sync window
            if not start_date or not end_date:
                start_date, end_date = get_sync_window()

            logger.info(f"Starting sync from {start_date.date()} to {end_date.date()}")

            # Read source events
            logger.info("Reading source events...")
            source_events = self.source_reader.read_events(
                calendar_id=source_calendar_id,
                start_date=start_date,
                end_date=end_date,
            )
            result.events_read = len(source_events)
            logger.info(f"Read {result.events_read} source events")

            # If read-only mode, we're done
            if not self.target_writer or dry_run:
                logger.info("Read-only mode or dry run - no changes made")
                return result

            # Process each event
            for event in source_events:
                try:
                    if not self.strategy.should_sync(event):
                        result.events_skipped += 1
                        continue

                    # For now, always create (no deduplication yet)
                    # In future: check if event exists in target
                    self.target_writer.create_event(
                        event,
                        calendar_id=target_calendar_id,
                    )
                    result.events_created += 1

                except Exception as e:
                    error_msg = f"Failed to sync event {event.subject}: {e}"
                    logger.error(error_msg)
                    result.errors.append(error_msg)

            logger.info(
                f"Sync complete: {result.events_created} created, "
                f"{result.events_updated} updated, "
                f"{result.events_skipped} skipped, "
                f"{len(result.errors)} errors"
            )

            return result

        except Exception as e:
            error_msg = f"Sync failed: {e}"
            logger.error(error_msg)
            result.errors.append(error_msg)
            return result

    def preview_sync(
        self,
        source_calendar_id: Optional[str] = None,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
    ) -> list:
        """
        Preview events that would be synced.

        Args:
            source_calendar_id: Source calendar ID
            start_date: Start date for preview
            end_date: End date for preview

        Returns:
            List of events that would be synced
        """
        if not start_date or not end_date:
            start_date, end_date = get_sync_window()

        events = self.source_reader.read_events(
            calendar_id=source_calendar_id,
            start_date=start_date,
            end_date=end_date,
        )

        return [e for e in events if self.strategy.should_sync(e)]
