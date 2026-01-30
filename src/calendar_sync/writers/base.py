"""Abstract base class for calendar writers."""

from abc import ABC, abstractmethod
from typing import Optional

from ..models.event import CalendarEvent


class CalendarWriter(ABC):
    """Abstract base class for calendar writers."""

    @abstractmethod
    def create_event(
        self,
        event: CalendarEvent,
        calendar_id: Optional[str] = None,
    ) -> str:
        """
        Create a new event.

        Args:
            event: CalendarEvent to create
            calendar_id: Calendar ID (None for default calendar)

        Returns:
            Created event ID

        Raises:
            CalendarWriteError: If event creation fails
        """

    @abstractmethod
    def update_event(
        self,
        event: CalendarEvent,
        calendar_id: Optional[str] = None,
    ) -> None:
        """
        Update an existing event.

        Args:
            event: CalendarEvent with updated data
            calendar_id: Calendar ID (None for default calendar)

        Raises:
            CalendarWriteError: If event update fails
        """

    @abstractmethod
    def delete_event(
        self,
        event_id: str,
        calendar_id: Optional[str] = None,
    ) -> None:
        """
        Delete an event.

        Args:
            event_id: Event identifier
            calendar_id: Calendar ID (None for default calendar)

        Raises:
            CalendarWriteError: If event deletion fails
        """
