"""Microsoft 365 calendar reader using Graph API."""

import logging
from datetime import datetime
from typing import Optional

from office365.graph_client import GraphClient

from ..auth.msal_auth import M365AuthProvider
from ..models.calendar import Calendar
from ..models.event import Attendee, CalendarEvent, EventStatus, Location
from ..utils.date_utils import ensure_utc
from ..utils.exceptions import CalendarReadError
from .base import CalendarReader

logger = logging.getLogger(__name__)


class M365CalendarReader(CalendarReader):
    """Read calendars from Microsoft 365 using Graph API."""

    def __init__(self, auth_provider: M365AuthProvider):
        """
        Initialize M365 calendar reader.

        Args:
            auth_provider: Microsoft 365 authentication provider
        """
        self.auth_provider = auth_provider
        self._client: Optional[GraphClient] = None

    @property
    def client(self) -> GraphClient:
        """Lazy-load Graph client."""
        if self._client is None:

            def token_func() -> dict[str, str]:
                return {"access_token": self.auth_provider.get_access_token()}

            self._client = GraphClient(token_func)
        return self._client

    def list_calendars(self) -> list[Calendar]:
        """List all calendars for the authenticated user."""
        try:
            calendars_result = self.client.me.calendars.get().execute_query()
            result = []

            for cal in calendars_result:
                result.append(
                    Calendar(
                        id=cal.id,
                        name=cal.name,
                        owner_email=(
                            cal.owner.address
                            if hasattr(cal, "owner") and hasattr(cal.owner, "address")
                            else None
                        ),
                        source_system="m365",
                        is_default=(
                            cal.is_default_calendar
                            if hasattr(cal, "is_default_calendar")
                            else False
                        ),
                        can_edit=cal.can_edit if hasattr(cal, "can_edit") else False,
                        color=cal.color if hasattr(cal, "color") else None,
                    )
                )

            logger.info(f"Found {len(result)} M365 calendars")
            return result

        except Exception as e:
            raise CalendarReadError(f"Failed to list M365 calendars: {e}") from e

    def read_events(
        self,
        calendar_id: Optional[str] = None,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
    ) -> list[CalendarEvent]:
        """Read events from M365 calendar."""
        try:
            # Get calendar reference
            if calendar_id:
                calendar = self.client.me.calendars[calendar_id]
            else:
                calendar = self.client.me.calendar

            # Build query
            query = calendar.events

            # Apply date filters using filter query
            if start_date or end_date:
                filter_parts = []
                if start_date:
                    # Graph API requires datetime without timezone info in filter
                    start_str = ensure_utc(start_date).strftime("%Y-%m-%dT%H:%M:%S")
                    filter_parts.append(f"start/dateTime ge '{start_str}'")
                if end_date:
                    # Graph API requires datetime without timezone info in filter
                    end_str = ensure_utc(end_date).strftime("%Y-%m-%dT%H:%M:%S")
                    filter_parts.append(f"end/dateTime le '{end_str}'")

                if filter_parts:
                    query = query.filter(" and ".join(filter_parts))

            # Execute query
            events_result = query.get().execute_query()

            # Transform to normalized model
            result = []
            for event in events_result:
                normalized = self._transform_event(event)
                result.append(normalized)

            logger.info(f"Read {len(result)} events from M365")
            return result

        except Exception as e:
            raise CalendarReadError(f"Failed to read M365 events: {e}") from e

    def get_event(
        self, event_id: str, calendar_id: Optional[str] = None
    ) -> CalendarEvent:
        """Get specific event from M365."""
        try:
            if calendar_id:
                event = (
                    self.client.me.calendars[calendar_id]
                    .events[event_id]
                    .get()
                    .execute_query()
                )
            else:
                event = (
                    self.client.me.calendar.events[event_id].get().execute_query()
                )

            return self._transform_event(event)

        except Exception as e:
            raise CalendarReadError(f"Failed to get M365 event {event_id}: {e}") from e

    def _transform_event(self, graph_event) -> CalendarEvent:
        """Transform Graph API event to normalized model."""
        # Parse attendees
        attendees = []
        if hasattr(graph_event, "attendees") and graph_event.attendees:
            for attendee in graph_event.attendees:
                attendees.append(
                    Attendee(
                        email=(
                            attendee.email_address.address
                            if hasattr(attendee, "email_address")
                            else ""
                        ),
                        name=(
                            attendee.email_address.name
                            if hasattr(attendee, "email_address")
                            and hasattr(attendee.email_address, "name")
                            else None
                        ),
                        response_status=(
                            attendee.status.response.lower()
                            if hasattr(attendee, "status")
                            and hasattr(attendee.status, "response")
                            else None
                        ),
                        is_organizer=False,
                    )
                )

        # Parse organizer
        organizer = None
        if hasattr(graph_event, "organizer") and graph_event.organizer:
            organizer = Attendee(
                email=(
                    graph_event.organizer.email_address.address
                    if hasattr(graph_event.organizer, "email_address")
                    else ""
                ),
                name=(
                    graph_event.organizer.email_address.name
                    if hasattr(graph_event.organizer, "email_address")
                    and hasattr(graph_event.organizer.email_address, "name")
                    else None
                ),
                is_organizer=True,
            )

        # Parse location
        location = None
        if hasattr(graph_event, "location") and graph_event.location:
            # Location can be a string or an object
            if isinstance(graph_event.location, str):
                location = Location(display_name=graph_event.location)
            elif hasattr(graph_event.location, "display_name"):
                # Extract address if it's an object
                address_str = None
                if hasattr(graph_event.location, "address"):
                    addr_obj = graph_event.location.address
                    if isinstance(addr_obj, str):
                        address_str = addr_obj
                    elif hasattr(addr_obj, "street"):
                        # Build address from PhysicalAddress object
                        parts = []
                        for attr in ["street", "city", "state", "postalCode", "countryOrRegion"]:
                            if hasattr(addr_obj, attr) and getattr(addr_obj, attr):
                                parts.append(str(getattr(addr_obj, attr)))
                        address_str = ", ".join(parts) if parts else None

                location = Location(
                    display_name=graph_event.location.display_name or "",
                    address=address_str,
                )
            elif hasattr(graph_event.location, "displayName"):
                # Handle different casing
                address_str = None
                if hasattr(graph_event.location, "address"):
                    addr_obj = graph_event.location.address
                    if isinstance(addr_obj, str):
                        address_str = addr_obj
                    elif hasattr(addr_obj, "street"):
                        parts = []
                        for attr in ["street", "city", "state", "postalCode", "countryOrRegion"]:
                            if hasattr(addr_obj, attr) and getattr(addr_obj, attr):
                                parts.append(str(getattr(addr_obj, attr)))
                        address_str = ", ".join(parts) if parts else None

                location = Location(
                    display_name=graph_event.location.displayName or "",
                    address=address_str,
                )

        # Parse dates - handle different attribute names
        start_str = getattr(graph_event.start, "dateTime", None) or getattr(
            graph_event.start, "date_time", None
        )
        end_str = getattr(graph_event.end, "dateTime", None) or getattr(
            graph_event.end, "date_time", None
        )

        start_dt = datetime.fromisoformat(start_str.replace("Z", "+00:00"))
        end_dt = datetime.fromisoformat(end_str.replace("Z", "+00:00"))

        # Determine status
        status = EventStatus.CONFIRMED
        if hasattr(graph_event, "is_cancelled") and graph_event.is_cancelled:
            status = EventStatus.CANCELLED
        elif hasattr(graph_event, "response_status"):
            if (
                hasattr(graph_event.response_status, "response")
                and graph_event.response_status.response == "tentativelyAccepted"
            ):
                status = EventStatus.TENTATIVE

        return CalendarEvent(
            id=graph_event.id,
            source_system="m365",
            ical_uid=getattr(graph_event, "ical_uid", None),
            subject=graph_event.subject or "(No Subject)",
            body=(
                getattr(graph_event.body, "content", None)
                if hasattr(graph_event, "body")
                else None
            ),
            body_preview=getattr(graph_event, "body_preview", None),
            start=ensure_utc(start_dt),
            end=ensure_utc(end_dt),
            is_all_day=graph_event.is_all_day,
            timezone=(
                graph_event.start.time_zone
                if hasattr(graph_event.start, "time_zone")
                else "UTC"
            ),
            organizer=organizer,
            attendees=attendees,
            location=location,
            status=status,
            sensitivity=getattr(graph_event, "sensitivity", "normal").lower(),
            show_as=getattr(graph_event, "show_as", "busy").lower(),
            categories=list(getattr(graph_event, "categories", [])),
            is_recurring=(
                hasattr(graph_event, "recurrence")
                and graph_event.recurrence is not None
            ),
            created=(
                datetime.fromisoformat(
                    graph_event.created_date_time.replace("Z", "+00:00")
                )
                if hasattr(graph_event, "created_date_time")
                and graph_event.created_date_time
                else None
            ),
            last_modified=(
                datetime.fromisoformat(
                    graph_event.last_modified_date_time.replace("Z", "+00:00")
                )
                if hasattr(graph_event, "last_modified_date_time")
                and graph_event.last_modified_date_time
                else None
            ),
        )
