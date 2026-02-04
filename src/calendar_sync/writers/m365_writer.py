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


# M365 preset color names -> API values
# See https://learn.microsoft.com/en-us/graph/api/resources/outlookcategory
CATEGORY_COLORS = {
    "red": "preset0",
    "orange": "preset1",
    "brown": "preset2",
    "yellow": "preset3",
    "green": "preset4",
    "teal": "preset5",
    "olive": "preset6",
    "blue": "preset7",
    "purple": "preset8",
    "cranberry": "preset9",
    "steel": "preset10",
    "darksteel": "preset11",
    "gray": "preset12",
    "darkgray": "preset13",
    "black": "preset14",
    "darkred": "preset15",
    "darkorange": "preset16",
    "darkyellow": "preset17",
    "darkgreen": "preset18",
    "darkteal": "preset19",
    "darkolive": "preset20",
    "darkblue": "preset21",
    "darkpurple": "preset22",
    "darkcranberry": "preset23",
    "none": "none",
}


class M365CalendarWriter(CalendarWriter):
    """Write events to Microsoft 365 using Graph API."""

    def __init__(self, auth_provider: M365AuthProvider, primary_email: Optional[str] = None):
        self.auth_provider = auth_provider
        self.primary_email = primary_email
        self.use_client_credentials = auth_provider.use_client_credentials
        self._existing_events: Optional[dict[str, str]] = None
        self._ensured_categories: set[str] = set()

        if self.use_client_credentials and not primary_email:
            raise CalendarWriteError(
                "primary_email is required when using client credentials flow (client_secret configured)"
            )

    @property
    def _user_path(self) -> str:
        """Get the correct user path based on auth type."""
        if self.use_client_credentials:
            # App-only auth: use /users/{email}
            return f"users/{self.primary_email}"
        else:
            # Delegated auth: use /me
            return "me"

    def _headers(self) -> dict[str, str]:
        token = self.auth_provider.get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    def ensure_category(self, name: str, color: str = "blue") -> None:
        """Ensure an Outlook category exists with the given color."""
        if name in self._ensured_categories:
            return

        preset = CATEGORY_COLORS.get(color.lower(), color)
        # Check if category already exists
        url = f"{GRAPH_BASE}/{self._user_path}/outlook/masterCategories"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        existing = {cat["displayName"]: cat for cat in resp.json().get("value", [])}

        if name in existing:
            # Update color if different
            cat = existing[name]
            if cat.get("color") != preset:
                patch_url = f"{url}/{cat['id']}"
                requests.patch(patch_url, headers=self._headers(), json={"color": preset})
                logger.info(f"Updated category '{name}' color to {color}")
        else:
            # Create new category
            resp = requests.post(url, headers=self._headers(), json={
                "displayName": name,
                "color": preset,
            })
            resp.raise_for_status()
            logger.info(f"Created category '{name}' with color {color}")

        self._ensured_categories.add(name)

    def get_existing_events(
        self, start: datetime, end: datetime, calendar_id: Optional[str] = None
    ) -> dict[str, str]:
        """Fetch existing events and return a dict of (subject, start) -> event_id for dedup."""
        if calendar_id:
            url = f"{GRAPH_BASE}/{self._user_path}/calendars/{calendar_id}/calendarView"
        else:
            url = f"{GRAPH_BASE}/{self._user_path}/calendarView"

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
                url = f"{GRAPH_BASE}/{self._user_path}/calendars/{calendar_id}/events"
            else:
                url = f"{GRAPH_BASE}/{self._user_path}/calendar/events"

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
                url = f"{GRAPH_BASE}/{self._user_path}/calendars/{calendar_id}/events/{event.id}"
            else:
                url = f"{GRAPH_BASE}/{self._user_path}/calendar/events/{event.id}"

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
                url = f"{GRAPH_BASE}/{self._user_path}/calendars/{calendar_id}/events/{event_id}"
            else:
                url = f"{GRAPH_BASE}/{self._user_path}/calendar/events/{event_id}"

            resp = requests.delete(url, headers=self._headers())
            resp.raise_for_status()
            logger.info(f"Deleted event: {event_id}")

        except Exception as e:
            raise CalendarWriteError(f"Failed to delete M365 event {event_id}: {e}") from e
