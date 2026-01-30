"""Exchange calendar reader using OWA REST API with Selenium-based cookie authentication."""

import logging
from datetime import datetime, timedelta
from typing import Any, Optional

import requests

from ..auth.selenium_auth import SeleniumEWSAuth
from ..config import EWSConfig
from ..models.calendar import Calendar
from ..models.event import Attendee, CalendarEvent, EventStatus, Location
from ..utils.date_utils import ensure_utc
from ..utils.exceptions import CalendarReadError
from .base import CalendarReader

logger = logging.getLogger(__name__)

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36"

# Days of week mapping for recurrence patterns
DAYS_OF_WEEK = {
    "Sunday": 6, "Monday": 0, "Tuesday": 1, "Wednesday": 2,
    "Thursday": 3, "Friday": 4, "Saturday": 5,
}


def _generate_weekly_occurrences(
    first_start: datetime,
    duration: timedelta,
    interval: int,
    days_of_week: list[int],
    rec_start: datetime,
    rec_end: datetime,
    window_start: datetime,
    window_end: datetime,
) -> list[tuple[datetime, datetime]]:
    """Generate weekly recurrence occurrences within a date window."""
    results = []
    # Start from the recurrence start date, aligned to the first day of the week
    current = rec_start.replace(hour=first_start.hour, minute=first_start.minute,
                                second=first_start.second, microsecond=0)
    # Go back to Monday of that week
    current -= timedelta(days=current.weekday())

    while current <= min(rec_end, window_end):
        for dow in sorted(days_of_week):
            day = current + timedelta(days=dow)
            if day < rec_start:
                continue
            if day > rec_end or day > window_end:
                break
            occ_start = day.replace(hour=first_start.hour, minute=first_start.minute,
                                    second=first_start.second)
            occ_end = occ_start + duration
            if occ_end >= window_start and occ_start <= window_end:
                results.append((ensure_utc(occ_start), ensure_utc(occ_end)))
        current += timedelta(weeks=interval)

    return results


def _generate_daily_occurrences(
    first_start: datetime,
    duration: timedelta,
    interval: int,
    rec_start: datetime,
    rec_end: datetime,
    window_start: datetime,
    window_end: datetime,
) -> list[tuple[datetime, datetime]]:
    """Generate daily recurrence occurrences within a date window."""
    results = []
    current = rec_start.replace(hour=first_start.hour, minute=first_start.minute,
                                second=first_start.second, microsecond=0)
    while current <= min(rec_end, window_end):
        occ_end = current + duration
        if occ_end >= window_start and current <= window_end:
            results.append((ensure_utc(current), ensure_utc(occ_end)))
        current += timedelta(days=interval)
    return results


def _parse_deleted_dates(master: dict) -> set[str]:
    """Extract deleted occurrence dates from a recurring master."""
    deleted = set()
    deleted_occs = master.get("DeletedOccurrences", []) or []
    for occ in deleted_occs:
        start_str = occ.get("Start", "")
        if start_str:
            try:
                dt = datetime.fromisoformat(start_str)
                deleted.add(dt.strftime("%Y-%m-%d"))
            except (ValueError, AttributeError):
                pass
    return deleted


def _parse_recurrence_dates(
    master: dict,
    window_start: datetime,
    window_end: datetime,
) -> list[tuple[datetime, datetime]]:
    """Parse recurrence pattern and generate occurrence dates within window."""
    recurrence = master.get("Recurrence", {})
    pattern = recurrence.get("RecurrencePattern", {})
    range_info = recurrence.get("RecurrenceRange", {})
    pattern_type = pattern.get("__type", "")

    # Check LastOccurrence — if it ends before window, skip entirely
    last_occ = master.get("LastOccurrence", {})
    if last_occ:
        last_end_str = last_occ.get("End")
        if last_end_str:
            try:
                last_end = datetime.fromisoformat(last_end_str)
                if last_end.tzinfo is None:
                    last_end = ensure_utc(last_end)
                if last_end < window_start:
                    return []
            except (ValueError, AttributeError):
                pass

    # Collect deleted occurrence dates to exclude
    deleted_dates = _parse_deleted_dates(master)

    first_occ = master.get("FirstOccurrence", {})
    first_start_str = first_occ.get("Start") or master.get("Start")
    first_end_str = first_occ.get("End") or master.get("End")
    if not first_start_str:
        return []

    first_start = datetime.fromisoformat(first_start_str)
    first_end = datetime.fromisoformat(first_end_str) if first_end_str else first_start + timedelta(minutes=30)
    duration = first_end - first_start

    # Parse recurrence range
    rec_start_str = range_info.get("StartDate", "")
    rec_end_str = range_info.get("EndDate", "")
    try:
        rec_start = datetime.fromisoformat(rec_start_str.replace("+01:00", "+00:00").replace("+02:00", "+00:00"))
        if rec_start.tzinfo is None:
            rec_start = ensure_utc(rec_start)
    except (ValueError, AttributeError):
        rec_start = first_start

    range_type = range_info.get("__type", "")
    if "NoEnd" in range_type:
        rec_end = window_end + timedelta(days=1)
    elif rec_end_str:
        try:
            rec_end = datetime.fromisoformat(rec_end_str.replace("+01:00", "+00:00").replace("+02:00", "+00:00"))
            if rec_end.tzinfo is None:
                rec_end = ensure_utc(rec_end)
        except (ValueError, AttributeError):
            rec_end = window_end + timedelta(days=1)
    else:
        rec_end = window_end + timedelta(days=1)

    # Skip if recurrence range doesn't overlap with window
    if rec_end < window_start or rec_start > window_end:
        return []

    interval = pattern.get("Interval", 1)

    raw_results = []
    if "Weekly" in pattern_type:
        days_str = pattern.get("DaysOfWeek", "")
        days = [DAYS_OF_WEEK[d.strip()] for d in days_str.split() if d.strip() in DAYS_OF_WEEK]
        if not days:
            return []
        raw_results = _generate_weekly_occurrences(
            first_start, duration, interval, days,
            rec_start, rec_end, window_start, window_end,
        )
    elif "Daily" in pattern_type:
        raw_results = _generate_daily_occurrences(
            first_start, duration, interval,
            rec_start, rec_end, window_start, window_end,
        )
    elif "Monthly" in pattern_type or "Yearly" in pattern_type:
        day_of_month = pattern.get("DayOfMonth", first_start.day)
        if "Yearly" in pattern_type:
            month = pattern.get("Month") or first_start.month
            if isinstance(month, str):
                months_map = {"January": 1, "February": 2, "March": 3, "April": 4,
                              "May": 5, "June": 6, "July": 7, "August": 8,
                              "September": 9, "October": 10, "November": 11, "December": 12}
                month = months_map.get(month, first_start.month)
            for year in range(window_start.year, window_end.year + 1):
                try:
                    occ = first_start.replace(year=year, month=month, day=day_of_month)
                    occ_end = occ + duration
                    if rec_start <= occ <= rec_end and occ_end >= window_start and occ <= window_end:
                        raw_results.append((ensure_utc(occ), ensure_utc(occ_end)))
                except ValueError:
                    pass
        else:
            current = window_start.replace(day=1)
            while current <= window_end:
                try:
                    occ = first_start.replace(year=current.year, month=current.month, day=day_of_month)
                    occ_end = occ + duration
                    if rec_start <= occ <= rec_end and occ_end >= window_start and occ <= window_end:
                        raw_results.append((ensure_utc(occ), ensure_utc(occ_end)))
                except ValueError:
                    pass
                if current.month == 12:
                    current = current.replace(year=current.year + 1, month=1)
                else:
                    current = current.replace(month=current.month + 1)
    else:
        logger.warning(f"Unsupported recurrence pattern: {pattern_type}")
        return []

    # Filter out deleted occurrences
    if deleted_dates:
        raw_results = [
            (s, e) for s, e in raw_results
            if s.strftime("%Y-%m-%d") not in deleted_dates
        ]

    return raw_results


class EWSSeleniumReader(CalendarReader):
    """Read calendars from Exchange using OWA REST API with cookie authentication."""

    def __init__(self, selenium_auth: SeleniumEWSAuth, config: EWSConfig):
        self.selenium_auth = selenium_auth
        self.config = config
        self._session: Optional[requests.Session] = None

    def _get_session(self) -> requests.Session:
        """Get or create an authenticated requests session."""
        if self._session is None:
            cookies = self.selenium_auth.get_cookies()
            self._session = requests.Session()
            self._session.cookies.update(cookies)
            self._session.headers.update({
                "User-Agent": USER_AGENT,
                "Content-Type": "application/json; charset=utf-8",
                "X-OWA-CANARY": cookies.get("X-OWA-CANARY", ""),
            })
        return self._session

    def _owa_action(self, action: str, body: dict) -> dict:
        """Call OWA service.svc with a given action and body."""
        session = self._get_session()
        url = f"{self.config.server_url.rstrip('/')}/owa/service.svc?action={action}"

        payload = {
            "__type": f"{action}JsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "Exchange2013",
            },
            "Body": body,
        }

        action_headers = {"Action": action}
        resp = session.post(url, json=payload, headers=action_headers, timeout=30)

        if resp.status_code in (401, 302, 403):
            logger.warning("OWA returned %s, refreshing cookies...", resp.status_code)
            self._session = None
            self.selenium_auth.get_cookies(force_refresh=True)
            session = self._get_session()
            resp = session.post(url, json=payload, headers=action_headers, timeout=30)

        if resp.status_code != 200:
            raise CalendarReadError(
                f"OWA {action} failed with status {resp.status_code}: {resp.text[:500]}"
            )

        data = resp.json()
        resp_body = data.get("Body", {})
        messages = resp_body.get("ResponseMessages", {}).get("Items", [])

        if messages and messages[0].get("ResponseClass") != "Success":
            code = messages[0].get("ResponseCode", "Unknown")
            raise CalendarReadError(f"OWA {action} error: {code}")

        return data

    def list_calendars(self) -> list[Calendar]:
        """List calendars from Exchange."""
        try:
            body = {
                "__type": "GetFolderRequest:#Exchange",
                "FolderShape": {"__type": "FolderResponseShape:#Exchange", "BaseShape": "Default"},
                "FolderIds": [{"__type": "DistinguishedFolderId:#Exchange", "Id": "calendar"}],
            }

            data = self._owa_action("GetFolder", body)
            folders = data["Body"]["ResponseMessages"]["Items"]
            name = "Calendar"
            if folders:
                folder = folders[0].get("Folders", [{}])
                if folder:
                    name = folder[0].get("DisplayName", "Calendar")

            result = [
                Calendar(
                    id="primary",
                    name=name,
                    owner_email=self.config.primary_email,
                    source_system="ews",
                    is_default=True,
                    can_edit=True,
                )
            ]

            logger.info(f"Found {len(result)} EWS calendars")
            return result

        except CalendarReadError:
            raise
        except Exception as e:
            raise CalendarReadError(f"Failed to list EWS calendars: {e}") from e

    def read_events(
        self,
        calendar_id: Optional[str] = None,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
    ) -> list[CalendarEvent]:
        """Read events from Exchange calendar."""
        try:
            start = ensure_utc(start_date) if start_date else ensure_utc(datetime(2020, 1, 1))
            end = ensure_utc(end_date) if end_date else ensure_utc(datetime(2030, 1, 1))

            # Fetch all items (OWA JSON API doesn't support CalendarView filtering)
            body = {
                "__type": "FindItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "Default",
                },
                "ParentFolderIds": [
                    {"__type": "DistinguishedFolderId:#Exchange", "Id": "calendar"}
                ],
                "Traversal": "Shallow",
            }

            data = self._owa_action("FindItem", body)
            items_container = data["Body"]["ResponseMessages"]["Items"][0]
            items = items_container.get("RootFolder", {}).get("Items", [])

            # Separate single events and recurring masters
            single_items = []
            master_items = []
            for item in items:
                if item.get("CalendarItemType") == "RecurringMaster":
                    master_items.append(item)
                else:
                    single_items.append(item)

            # Filter single events by date
            result = []
            for item in single_items:
                event = self._parse_item(item)
                if event.end >= start and event.start <= end:
                    if self._should_skip(event):
                        continue
                    result.append(event)

            # Expand recurring masters
            if master_items:
                result.extend(self._expand_recurring_masters(master_items, start, end))

            result.sort(key=lambda e: e.start)
            logger.info(
                f"Read {len(result)} events from EWS "
                f"({len(single_items)} single, {len(master_items)} recurring masters)"
            )
            return result

        except CalendarReadError:
            raise
        except Exception as e:
            raise CalendarReadError(f"Failed to read EWS events: {e}") from e

    def _expand_recurring_masters(
        self,
        masters: list[dict],
        window_start: datetime,
        window_end: datetime,
    ) -> list[CalendarEvent]:
        """Expand recurring master items into individual occurrences."""
        # Batch-fetch recurrence details for all masters
        BATCH_SIZE = 50
        detailed_masters = []

        for i in range(0, len(masters), BATCH_SIZE):
            batch = masters[i:i + BATCH_SIZE]
            item_ids = [
                {
                    "__type": "ItemId:#Exchange",
                    "Id": m["ItemId"]["Id"],
                    "ChangeKey": m["ItemId"]["ChangeKey"],
                }
                for m in batch
            ]

            body = {
                "__type": "GetItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "AllProperties",
                },
                "ItemIds": item_ids,
            }

            data = self._owa_action("GetItem", body)
            for msg in data["Body"]["ResponseMessages"]["Items"]:
                if msg.get("ResponseClass") == "Success" and msg.get("Items"):
                    detailed_masters.append(msg["Items"][0])

        # Generate occurrences from recurrence patterns
        occurrences = []
        for master in detailed_masters:
            base_event = self._parse_item(master)
            if self._should_skip(base_event):
                continue

            last_occ = master.get("LastOccurrence", {})
            logger.debug(
                f"Recurring master: '{base_event.subject}' | "
                f"LastOccurrence.End={last_occ.get('End') if last_occ else 'N/A'} | "
                f"DeletedOccurrences={len(master.get('DeletedOccurrences', []) or [])} | "
                f"Recurrence.Range={master.get('Recurrence', {}).get('RecurrenceRange', {})}"
            )

            dates = _parse_recurrence_dates(master, window_start, window_end)
            if dates:
                logger.debug(f"  -> Generated {len(dates)} occurrences for '{base_event.subject}'")
            for occ_start, occ_end in dates:
                event = self._parse_item(master)
                event.start = occ_start
                event.end = occ_end
                event.is_recurring = True
                occurrences.append(event)

        logger.info(
            f"Expanded {len(detailed_masters)} recurring masters into {len(occurrences)} occurrences"
        )
        return occurrences

    def get_event(
        self, event_id: str, calendar_id: Optional[str] = None
    ) -> CalendarEvent:
        """Get specific event from Exchange."""
        try:
            body = {
                "__type": "GetItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "AllProperties",
                },
                "ItemIds": [{"__type": "ItemId:#Exchange", "Id": event_id}],
            }

            data = self._owa_action("GetItem", body)
            items = data["Body"]["ResponseMessages"]["Items"][0].get("Items", [])
            if not items:
                raise CalendarReadError(f"Event {event_id} not found")
            return self._parse_item(items[0])

        except CalendarReadError:
            raise
        except Exception as e:
            raise CalendarReadError(f"Failed to get EWS event {event_id}: {e}") from e

    @staticmethod
    def _should_skip(event: CalendarEvent) -> bool:
        """Check if an event should be excluded from results."""
        subject = (event.subject or "").strip().lower()
        if subject == "out of office":
            return True
        if subject.startswith("canceled:") or subject.startswith("cancelled:"):
            return True
        if event.status == EventStatus.CANCELLED:
            return True
        return False

    def _parse_item(self, item: dict[str, Any]) -> CalendarEvent:
        """Parse an OWA CalendarItem JSON into a CalendarEvent."""
        item_id = item.get("ItemId", {}).get("Id", "")
        subject = item.get("Subject") or "(No Subject)"
        body_text = None
        body_obj = item.get("Body")
        if isinstance(body_obj, dict):
            body_text = body_obj.get("Value")
        elif isinstance(body_obj, str):
            body_text = body_obj

        start_str = item.get("Start")
        end_str = item.get("End")
        start_dt = datetime.fromisoformat(start_str) if start_str else datetime.now()
        end_dt = datetime.fromisoformat(end_str) if end_str else start_dt

        location_str = item.get("Location")
        if isinstance(location_str, dict):
            location_str = location_str.get("DisplayName")
        location = Location(display_name=location_str) if location_str else None

        is_all_day = bool(item.get("IsAllDayEvent", False))
        is_cancelled = bool(item.get("IsCancelled", False))
        is_recurring = item.get("CalendarItemType") in ("RecurringMaster", "Occurrence")
        sensitivity = (item.get("Sensitivity") or "Normal").lower()
        show_as = (item.get("LegacyFreeBusyStatus") or "Busy").lower()
        uid = item.get("UID")
        date_created = item.get("DateTimeCreated")
        date_modified = item.get("LastModifiedTime")

        # Parse organizer
        organizer = None
        org = item.get("Organizer")
        if org:
            mailbox = org.get("Mailbox", org)
            organizer = Attendee(
                email=mailbox.get("EmailAddress", ""),
                name=mailbox.get("Name"),
                is_organizer=True,
            )

        # Parse attendees
        attendees = []
        for att in item.get("RequiredAttendees", []) or []:
            mailbox = att.get("Mailbox", att)
            attendees.append(
                Attendee(
                    email=mailbox.get("EmailAddress", ""),
                    name=mailbox.get("Name"),
                    is_organizer=False,
                )
            )

        status = EventStatus.CANCELLED if is_cancelled else EventStatus.CONFIRMED

        return CalendarEvent(
            id=item_id,
            source_system="ews",
            ical_uid=uid,
            subject=subject,
            body=body_text,
            start=ensure_utc(start_dt),
            end=ensure_utc(end_dt),
            is_all_day=is_all_day,
            timezone="UTC",
            organizer=organizer,
            attendees=attendees,
            location=location,
            status=status,
            sensitivity=sensitivity,
            show_as=show_as,
            categories=item.get("Categories", []) or [],
            is_recurring=is_recurring,
            created=datetime.fromisoformat(date_created) if date_created else None,
            last_modified=datetime.fromisoformat(date_modified) if date_modified else None,
        )
