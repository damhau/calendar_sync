"""Exchange EWS calendar reader."""

import logging
from datetime import datetime
from typing import Optional

from exchangelib import DELEGATE, Account, Configuration
from exchangelib.items import CalendarItem

from ..auth.ews_auth import EWSAuthProvider
from ..config import EWSConfig
from ..models.calendar import Calendar
from ..models.event import Attendee, CalendarEvent, EventStatus, Location
from ..utils.date_utils import ensure_utc
from ..utils.exceptions import CalendarReadError
from .base import CalendarReader

logger = logging.getLogger(__name__)


class EWSCalendarReader(CalendarReader):
    """Read calendars from Exchange EWS."""

    def __init__(self, auth_provider: EWSAuthProvider, config: EWSConfig):
        """
        Initialize EWS calendar reader.

        Args:
            auth_provider: Exchange EWS authentication provider
            config: Exchange EWS configuration
        """
        self.auth_provider = auth_provider
        self.config = config
        self._account: Optional[Account] = None

    @property
    def account(self) -> Account:
        """Lazy-load EWS account."""
        if self._account is None:
            credentials = self.auth_provider.get_credentials()

            exchange_config = Configuration(
                server=self.config.server_url,
                credentials=credentials,
                auth_type="OAuth2",
            )

            self._account = Account(
                primary_smtp_address=self.config.primary_email,
                config=exchange_config,
                autodiscover=False,
                access_type=DELEGATE,
            )

        return self._account

    def list_calendars(self) -> list[Calendar]:
        """List calendars from EWS."""
        try:
            # EWS typically has one main calendar per mailbox
            calendar = self.account.calendar

            result = [
                Calendar(
                    id="primary",
                    name="Calendar",
                    owner_email=self.config.primary_email,
                    source_system="ews",
                    is_default=True,
                    can_edit=True,
                )
            ]

            logger.info(f"Found {len(result)} EWS calendars")
            return result

        except Exception as e:
            raise CalendarReadError(f"Failed to list EWS calendars: {e}") from e

    def read_events(
        self,
        calendar_id: Optional[str] = None,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
    ) -> list[CalendarEvent]:
        """Read events from EWS calendar."""
        try:
            calendar = self.account.calendar

            # Build filter
            query = calendar.filter()

            if start_date:
                query = query.filter(end__gte=ensure_utc(start_date))
            if end_date:
                query = query.filter(start__lte=ensure_utc(end_date))

            # Fetch events
            events = list(query)

            # Transform to normalized model
            result = []
            for event in events:
                normalized = self._transform_event(event)
                result.append(normalized)

            logger.info(f"Read {len(result)} events from EWS")
            return result

        except Exception as e:
            raise CalendarReadError(f"Failed to read EWS events: {e}") from e

    def get_event(
        self, event_id: str, calendar_id: Optional[str] = None
    ) -> CalendarEvent:
        """Get specific event from EWS."""
        try:
            # EWS uses item IDs
            calendar = self.account.calendar
            event = calendar.get(id=event_id)
            return self._transform_event(event)

        except Exception as e:
            raise CalendarReadError(f"Failed to get EWS event {event_id}: {e}") from e

    def _transform_event(self, ews_event: CalendarItem) -> CalendarEvent:
        """Transform EWS CalendarItem to normalized model."""
        # Parse attendees
        attendees = []
        if ews_event.required_attendees:
            for attendee in ews_event.required_attendees:
                attendees.append(
                    Attendee(
                        email=attendee.mailbox.email_address,
                        name=attendee.mailbox.name if attendee.mailbox.name else None,
                        response_status=(
                            attendee.response_type.lower()
                            if hasattr(attendee, "response_type")
                            else None
                        ),
                        is_organizer=False,
                    )
                )

        # Parse organizer
        organizer = None
        if ews_event.organizer:
            organizer = Attendee(
                email=ews_event.organizer.email_address,
                name=ews_event.organizer.name if ews_event.organizer.name else None,
                is_organizer=True,
            )

        # Parse location
        location = None
        if ews_event.location:
            location = Location(display_name=ews_event.location)

        # Determine status
        status = EventStatus.CONFIRMED
        if hasattr(ews_event, "is_cancelled") and ews_event.is_cancelled:
            status = EventStatus.CANCELLED

        return CalendarEvent(
            id=str(ews_event.id),
            source_system="ews",
            ical_uid=getattr(ews_event, "uid", None),
            subject=ews_event.subject or "(No Subject)",
            body=(
                ews_event.text_body
                if hasattr(ews_event, "text_body") and ews_event.text_body
                else None
            ),
            start=ensure_utc(ews_event.start),
            end=ensure_utc(ews_event.end),
            is_all_day=(
                ews_event.is_all_day
                if hasattr(ews_event, "is_all_day")
                else False
            ),
            timezone="UTC",
            organizer=organizer,
            attendees=attendees,
            location=location,
            status=status,
            sensitivity=(
                ews_event.sensitivity.lower()
                if hasattr(ews_event, "sensitivity")
                else "normal"
            ),
            show_as=(
                ews_event.legacy_free_busy_status.lower()
                if hasattr(ews_event, "legacy_free_busy_status")
                else "busy"
            ),
            categories=(
                list(ews_event.categories)
                if hasattr(ews_event, "categories") and ews_event.categories
                else []
            ),
            is_recurring=(
                ews_event.is_recurring
                if hasattr(ews_event, "is_recurring")
                else False
            ),
            created=(
                ensure_utc(ews_event.datetime_created)
                if hasattr(ews_event, "datetime_created")
                and ews_event.datetime_created
                else None
            ),
            last_modified=(
                ensure_utc(ews_event.last_modified_time)
                if hasattr(ews_event, "last_modified_time")
                and ews_event.last_modified_time
                else None
            ),
        )
