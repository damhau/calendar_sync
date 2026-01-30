"""Normalized calendar event data model."""

from datetime import datetime
from enum import Enum
from typing import Optional

import pytz
from pydantic import BaseModel, Field


class EventStatus(str, Enum):
    """Event status enumeration."""

    CONFIRMED = "confirmed"
    TENTATIVE = "tentative"
    CANCELLED = "cancelled"


class Attendee(BaseModel):
    """Event attendee."""

    email: str
    name: Optional[str] = None
    response_status: Optional[str] = None  # accepted, declined, tentative, none
    is_organizer: bool = False


class Location(BaseModel):
    """Event location."""

    display_name: str
    address: Optional[str] = None
    coordinates: Optional[tuple[float, float]] = None


class EventRecurrence(BaseModel):
    """Event recurrence pattern."""

    pattern: str  # daily, weekly, monthly, yearly
    interval: int = 1
    days_of_week: Optional[list[str]] = None
    day_of_month: Optional[int] = None
    month_of_year: Optional[int] = None
    end_date: Optional[datetime] = None
    occurrences: Optional[int] = None


class CalendarEvent(BaseModel):
    """Normalized calendar event model."""

    # Identifiers
    id: str  # Source system ID
    source_system: str  # "ews" or "m365"
    ical_uid: Optional[str] = None  # iCalendar UID for matching

    # Basic properties
    subject: str
    body: Optional[str] = None
    body_preview: Optional[str] = None

    # Time properties
    start: datetime
    end: datetime
    is_all_day: bool = False
    timezone: str = "UTC"

    # People
    organizer: Optional[Attendee] = None
    attendees: list[Attendee] = Field(default_factory=list)

    # Location
    location: Optional[Location] = None

    # Status and classification
    status: EventStatus = EventStatus.CONFIRMED
    sensitivity: str = "normal"  # normal, personal, private, confidential
    show_as: str = "busy"  # free, tentative, busy, oof, workingElsewhere

    # Categorization
    categories: list[str] = Field(default_factory=list)

    # Recurrence
    is_recurring: bool = False
    recurrence: Optional[EventRecurrence] = None
    recurrence_master_id: Optional[str] = None

    # Metadata
    created: Optional[datetime] = None
    last_modified: Optional[datetime] = None

    # Sync metadata
    sync_timestamp: datetime = Field(default_factory=lambda: datetime.now(pytz.utc))

    model_config = {
        "json_encoders": {
            datetime: lambda v: v.isoformat(),
        }
    }

    def to_local_time(self, target_tz: str = "UTC") -> "CalendarEvent":
        """
        Convert event times to target timezone.

        Args:
            target_tz: Target timezone name

        Returns:
            New CalendarEvent with converted times
        """
        tz = pytz.timezone(target_tz)
        event_copy = self.model_copy(deep=True)

        if self.start.tzinfo is None:
            event_copy.start = pytz.utc.localize(self.start).astimezone(tz)
        else:
            event_copy.start = self.start.astimezone(tz)

        if self.end.tzinfo is None:
            event_copy.end = pytz.utc.localize(self.end).astimezone(tz)
        else:
            event_copy.end = self.end.astimezone(tz)

        event_copy.timezone = target_tz
        return event_copy
