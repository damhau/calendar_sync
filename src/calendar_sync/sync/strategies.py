"""Sync strategies for calendar synchronization."""

from enum import Enum
from typing import Protocol

from ..models.event import CalendarEvent


class SyncDirection(str, Enum):
    """Sync direction options."""

    READ_ONLY = "read_only"
    WRITE_TO_M365 = "write_to_m365"
    BIDIRECTIONAL = "bidirectional"  # Future


class SyncStrategy(Protocol):
    """Protocol for sync strategies."""

    def should_sync(self, event: CalendarEvent) -> bool:
        """
        Determine if event should be synced.

        Args:
            event: CalendarEvent to evaluate

        Returns:
            True if event should be synced, False otherwise
        """
        ...

    def resolve_conflict(
        self,
        source_event: CalendarEvent,
        target_event: CalendarEvent,
    ) -> CalendarEvent:
        """
        Resolve conflicts between events.

        Args:
            source_event: Event from source system
            target_event: Event from target system

        Returns:
            Resolved CalendarEvent
        """
        ...


class OneWaySyncStrategy:
    """One-way sync strategy (source -> target)."""

    def should_sync(self, event: CalendarEvent) -> bool:
        """All events should be synced in one-way mode."""
        return True

    def resolve_conflict(
        self,
        source_event: CalendarEvent,
        target_event: CalendarEvent,
    ) -> CalendarEvent:
        """Source always wins in one-way sync."""
        return source_event
