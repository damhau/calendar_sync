"""Microsoft 365 calendar writer using Graph API directly."""

import logging
from datetime import datetime
from typing import Optional

import requests

from ..auth.msal_auth import M365AuthProvider
from ..models.event import CalendarEvent
from ..utils.exceptions import CalendarWriteError
from .base import CalendarWriter

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class M365CalendarWriter(CalendarWriter):
    """Write events to Microsoft 365 using Graph API."""

    def __init__(self, auth_provider: M365AuthProvider):
        self.auth_provider = auth_provider
        self._existing_events: Optional[dict[str, str]] = None

    def _headers(self) -> dict[str, str]:
        token = self.auth_provider.get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    def get_existing_events(
        self, start: datetime, end: datetime, calendar_id: Optional[str] = None
    ) -> dict[str, str]:
        """Fetch existing events and return a dict of (subject, start) -> event_id for dedup."""
        if calendar_id:
            url = f"{GRAPH_BASE}/me/calendars/{calendar_id}/calendarView"
        else:
            url = f"{GRAPH_BASE}/me/calendarView"

        params = {
            "startDateTime": start.strftime("%Y-%m-%dT%H:%M:%S"),
            "endDateTime": end.strftime("%Y-%m-%dT%H:%M:%S"),
            "$select": "id,subject,start",
            "$top": 500,
        }
        existing = {}
        while url:
            resp = requests.get(url, headers=self._headers(), params=params)
            resp.raise_for_status()
            data = resp.json()
            for ev in data.get("value", []):
                key = (ev.get("subject", ""), ev["start"].get("dateTime", "")[:16])
                existing[key] = ev["id"]
            url = data.get("@odata.nextLink")
            params = None  # nextLink includes params
        logger.info(f"Found {len(existing)} existing events in target calendar")
        return existing

    def _to_graph_format(self, event: CalendarEvent) -> dict:
        data = {
            "subject": event.subject,
            "start": {
                "dateTime": event.start.strftime("%Y-%m-%dT%H:%M:%S"),
                "timeZone": event.timezone,
            },
            "end": {
                "dateTime": event.end.strftime("%Y-%m-%dT%H:%M:%S"),
                "timeZone": event.timezone,
            },
            "isAllDay": event.is_all_day,
            "sensitivity": event.sensitivity,
            "showAs": event.show_as,
        }

        if event.body:
            data["body"] = {"contentType": "text", "content": event.body}

        if event.location:
            data["location"] = {"displayName": event.location.display_name}

        if event.categories:
            data["categories"] = event.categories

        return data

    def create_event(
        self,
        event: CalendarEvent,
        calendar_id: Optional[str] = None,
    ) -> str:
        try:
            if calendar_id:
                url = f"{GRAPH_BASE}/me/calendars/{calendar_id}/events"
            else:
                url = f"{GRAPH_BASE}/me/calendar/events"

            resp = requests.post(url, headers=self._headers(), json=self._to_graph_format(event))
            resp.raise_for_status()
            event_id = resp.json().get("id", "")
            logger.info(f"Created event: {event.subject}")
            return event_id

        except Exception as e:
            raise CalendarWriteError(f"Failed to create M365 event: {e}") from e

    def update_event(
        self,
        event: CalendarEvent,
        calendar_id: Optional[str] = None,
    ) -> None:
        try:
            if calendar_id:
                url = f"{GRAPH_BASE}/me/calendars/{calendar_id}/events/{event.id}"
            else:
                url = f"{GRAPH_BASE}/me/calendar/events/{event.id}"

            resp = requests.patch(url, headers=self._headers(), json=self._to_graph_format(event))
            resp.raise_for_status()
            logger.info(f"Updated event: {event.subject}")

        except Exception as e:
            raise CalendarWriteError(f"Failed to update M365 event {event.id}: {e}") from e

    def delete_event(
        self,
        event_id: str,
        calendar_id: Optional[str] = None,
    ) -> None:
        try:
            if calendar_id:
                url = f"{GRAPH_BASE}/me/calendars/{calendar_id}/events/{event_id}"
            else:
                url = f"{GRAPH_BASE}/me/calendar/events/{event_id}"

            resp = requests.delete(url, headers=self._headers())
            resp.raise_for_status()
            logger.info(f"Deleted event: {event_id}")

        except Exception as e:
            raise CalendarWriteError(f"Failed to delete M365 event {event_id}: {e}") from e
