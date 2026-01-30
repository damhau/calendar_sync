"""Abstract base class for calendar readers."""

from abc import ABC, abstractmethod
from datetime import datetime
from typing import Optional

from ..models.calendar import Calendar
from ..models.event import CalendarEvent


class CalendarReader(ABC):
    """Abstract base class for calendar readers."""

    @abstractmethod
    def list_calendars(self) -> list[Calendar]:
        """
        List all available calendars.

        Returns:
            List of Calendar objects

        Raises:
            CalendarReadError: If listing calendars fails
        """

    @abstractmethod
    def read_events(
        self,
        calendar_id: Optional[str] = None,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
    ) -> list[CalendarEvent]:
        """
        Read events from calendar(s).

        Args:
            calendar_id: Calendar ID (None for default calendar)
            start_date: Start date for event range
            end_date: End date for event range

        Returns:
            List of normalized CalendarEvent objects

        Raises:
            CalendarReadError: If reading events fails
        """

    @abstractmethod
    def get_event(
        self, event_id: str, calendar_id: Optional[str] = None
    ) -> CalendarEvent:
        """
        Get a specific event by ID.

        Args:
            event_id: Event identifier
            calendar_id: Calendar ID (None for default calendar)

        Returns:
            CalendarEvent object

        Raises:
            CalendarReadError: If getting event fails
        """
