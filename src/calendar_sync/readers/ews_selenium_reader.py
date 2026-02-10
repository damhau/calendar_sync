"""Exchange calendar reader using OWA REST API with Selenium-based cookie authentication."""

import json
import logging
import re
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
        self._canary_token: Optional[str] = None

    def _fetch_canary_token(self, session: requests.Session) -> Optional[str]:
        """Fetch X-OWA-CANARY token from OWA endpoints."""
        import re
        
        # Method 1: Load OWA page and extract canary from HTML/cookies
        try:
            url = f"{self.config.server_url.rstrip('/')}/owa/"
            resp = session.get(url, timeout=15, allow_redirects=True)
            
            # Check if canary was set as a cookie
            canary = session.cookies.get("X-OWA-CANARY")
            if canary:
                logger.info("✅ Fetched X-OWA-CANARY from cookie after OWA page load")
                return canary
            
            # Check response headers
            canary = resp.headers.get("X-OWA-CANARY")
            if canary:
                logger.info("✅ Fetched X-OWA-CANARY from response headers")
                return canary
            
            # Try to extract from page content
            if resp.status_code == 200:
                # Look for canary in various formats
                patterns = [
                    r'"canary"\s*:\s*"([^"]+)"',
                    r"'canary'\s*:\s*'([^']+)'",
                    r'X-OWA-CANARY["\s:]+["\']([^"\']+)["\']',
                    r'owaCanary["\s:]+["\']([^"\']+)["\']',
                ]
                for pattern in patterns:
                    match = re.search(pattern, resp.text)
                    if match:
                        logger.info(f"✅ Extracted X-OWA-CANARY from OWA page content")
                        return match.group(1)
        except Exception as e:
            logger.debug(f"Failed to get canary from OWA page: {e}")
        
        # Method 2: Try sessiondata.ashx
        try:
            url = f"{self.config.server_url.rstrip('/')}/owa/sessiondata.ashx?cid=0&fmt=json"
            resp = session.get(url, timeout=15)
            if resp.status_code == 200:
                try:
                    data = resp.json()
                    canary = data.get("owaCanary") or data.get("canary")
                    if canary:
                        logger.info("✅ Fetched X-OWA-CANARY from sessiondata.ashx")
                        return canary
                except Exception:
                    pass
        except Exception as e:
            logger.debug(f"Failed to get canary from sessiondata: {e}")
        
        # Method 3: Try the service endpoint which might return canary in response headers
        try:
            url = f"{self.config.server_url.rstrip('/')}/owa/service.svc?action=GetOwaUserConfiguration"
            headers = {
                "Content-Type": "application/json",
                "Action": "GetOwaUserConfiguration",
            }
            payload = {
                "__type": "GetOwaUserConfigurationRequest:#Exchange",
                "Header": {
                    "__type": "JsonRequestHeaders:#Exchange",
                    "RequestServerVersion": "Exchange2013"
                },
                "Body": {}
            }
            resp = session.post(url, json=payload, headers=headers, timeout=15)
            
            canary = resp.headers.get("X-OWA-CANARY")
            if canary:
                logger.info("✅ Fetched X-OWA-CANARY from service response headers")
                return canary
            
            # Check if set as cookie
            canary = session.cookies.get("X-OWA-CANARY")
            if canary:
                logger.info("✅ Fetched X-OWA-CANARY from cookie after service call")
                return canary
                    
        except Exception as e:
            logger.debug(f"Failed to fetch canary from service: {e}")
        
        return None

    def _get_session(self) -> requests.Session:
        """Get or create an authenticated requests session."""
        if self._session is None:
            cookies = self.selenium_auth.get_cookies()
            self._session = requests.Session()
            self._session.cookies.update(cookies)
            
            # Get canary token from cookies or fetch it
            self._canary_token = cookies.get("X-OWA-CANARY", "")
            
            if not self._canary_token:
                logger.warning("X-OWA-CANARY not in cookies, attempting to fetch...")
                self._canary_token = self._fetch_canary_token(self._session) or ""
            
            if not self._canary_token:
                logger.warning("⚠️  No X-OWA-CANARY token available - will try REST API fallback")
            else:
                logger.info(f"✅ Using X-OWA-CANARY token: {self._canary_token[:20]}...")
            
            self._session.headers.update({
                "User-Agent": USER_AGENT,
                "Content-Type": "application/json; charset=utf-8",
                "X-OWA-CANARY": self._canary_token,
            })
        return self._session

    def _is_office365(self) -> bool:
        """Check if this is Office 365 (outlook.office.com) vs on-premise Exchange."""
        return "outlook.office" in self.config.server_url.lower()

    def _rest_api_get_events(
        self,
        start: datetime,
        end: datetime,
    ) -> Optional[list[dict]]:
        """Try to get events using Outlook REST API v2.0 (works for Office 365)."""
        session = self._get_session()
        
        # Format dates for REST API
        start_str = start.strftime("%Y-%m-%dT%H:%M:%SZ")
        end_str = end.strftime("%Y-%m-%dT%H:%M:%SZ")
        
        # Try the Outlook REST API v2.0 calendarview endpoint
        url = f"{self.config.server_url.rstrip('/')}/api/v2.0/me/calendarview"
        params = {
            "startDateTime": start_str,
            "endDateTime": end_str,
            "$top": 500,
            "$select": "Id,Subject,Start,End,Location,Organizer,Attendees,IsAllDay,IsCancelled,ShowAs,Body,Categories,Recurrence",
        }
        
        try:
            resp = session.get(url, params=params, timeout=30)
            logger.warning(f"REST API v2.0 calendarview: status={resp.status_code}, url={url}")
            if resp.status_code != 200:
                logger.warning(f"REST API v2.0 error response: {resp.text[:500]}")
            
            if resp.status_code == 200:
                data = resp.json()
                events = data.get("value", [])
                logger.info(f"✅ REST API returned {len(events)} events")
                return events
        except Exception as e:
            logger.warning(f"REST API calendarview exception: {e}")
        
        # Try OWA calendar API endpoint used by the web UI
        try:
            url = f"{self.config.server_url.rstrip('/')}/owa/calendar/api/v1.0/me/calendarview"
            params = {
                "startDateTime": start_str,
                "endDateTime": end_str,
            }
            resp = session.get(url, params=params, timeout=30)
            logger.warning(f"OWA calendar API: status={resp.status_code}")
            if resp.status_code == 200:
                data = resp.json()
                events = data.get("value", [])
                logger.info(f"✅ OWA calendar API returned {len(events)} events")
                return events
            else:
                logger.warning(f"OWA calendar API error: {resp.text[:300]}")
        except Exception as e:
            logger.warning(f"OWA calendar API exception: {e}")

        # Try alternative endpoint: /owa/0/calendars/action/finditem
        try:
            url = f"{self.config.server_url.rstrip('/')}/owa/0/calendars/action/finditem"
            payload = {
                "request": {
                    "__type": "FindItemRequest:#Exchange",
                    "ItemShape": {"BaseShape": "Default"},
                    "ParentFolderIds": [{"__type": "DistinguishedFolderId:#Exchange", "Id": "calendar"}],
                    "Traversal": "Shallow",
                }
            }
            resp = session.post(url, json=payload, timeout=30)
            logger.warning(f"OWA finditem endpoint: status={resp.status_code}")
            
            if resp.status_code == 200:
                data = resp.json()
                logger.info(f"✅ OWA calendars endpoint returned data")
                return data.get("Items", [])
            else:
                logger.warning(f"OWA finditem error: {resp.text[:300]}")
        except Exception as e:
            logger.warning(f"OWA calendars endpoint failed: {e}")
        
        return None

    def _parse_rest_api_events(
        self,
        events: list[dict],
        start: datetime,
        end: datetime,
    ) -> list[CalendarEvent]:
        """Parse events from Outlook REST API v2.0 format."""
        results = []
        
        for item in events:
            try:
                # REST API uses different field names than OWA service
                item_id = item.get("Id", "")
                subject = item.get("Subject") or "(No Subject)"
                
                # Parse body
                body_obj = item.get("Body", {})
                body_text = body_obj.get("Content") if isinstance(body_obj, dict) else None
                
                # Parse dates - REST API format: {"DateTime": "...", "TimeZone": "..."}
                start_obj = item.get("Start", {})
                end_obj = item.get("End", {})
                
                if isinstance(start_obj, dict):
                    start_str = start_obj.get("DateTime", "")
                else:
                    start_str = str(start_obj) if start_obj else ""
                    
                if isinstance(end_obj, dict):
                    end_str = end_obj.get("DateTime", "")
                else:
                    end_str = str(end_obj) if end_obj else ""
                
                if not start_str:
                    continue
                    
                start_dt = datetime.fromisoformat(start_str.replace("Z", "+00:00"))
                end_dt = datetime.fromisoformat(end_str.replace("Z", "+00:00")) if end_str else start_dt + timedelta(hours=1)
                
                # Parse location
                loc_obj = item.get("Location", {})
                location = None
                if loc_obj:
                    loc_name = loc_obj.get("DisplayName") if isinstance(loc_obj, dict) else str(loc_obj)
                    if loc_name:
                        location = Location(display_name=loc_name)
                
                is_all_day = item.get("IsAllDay", False)
                is_cancelled = item.get("IsCancelled", False)
                is_recurring = item.get("Recurrence") is not None
                
                # Parse organizer
                organizer = None
                org = item.get("Organizer", {})
                if org:
                    email_addr = org.get("EmailAddress", {})
                    if email_addr:
                        organizer = Attendee(
                            email=email_addr.get("Address", ""),
                            name=email_addr.get("Name"),
                            is_organizer=True,
                        )
                
                # Parse attendees
                attendees = []
                for att in item.get("Attendees", []) or []:
                    email_addr = att.get("EmailAddress", {})
                    if email_addr:
                        attendees.append(
                            Attendee(
                                email=email_addr.get("Address", ""),
                                name=email_addr.get("Name"),
                                is_organizer=False,
                            )
                        )
                
                show_as = (item.get("ShowAs") or "Busy").lower()
                status = EventStatus.CANCELLED if is_cancelled else EventStatus.CONFIRMED
                
                event = CalendarEvent(
                    id=item_id,
                    source_system="ews",
                    ical_uid=item.get("iCalUId"),
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
                    sensitivity=(item.get("Sensitivity") or "Normal").lower(),
                    show_as=show_as,
                    categories=item.get("Categories", []) or [],
                    is_recurring=is_recurring,
                )
                
                if not self._should_skip(event):
                    results.append(event)
                    
            except Exception as e:
                logger.warning(f"Failed to parse REST API event: {e}")
                continue
        
        logger.info(f"Parsed {len(results)} events from REST API")
        return results

    def _parse_dom_events(
        self,
        events: list[dict],
        start: datetime,
        end: datetime,
    ) -> list[CalendarEvent]:
        """Parse events extracted from OWA DOM (aria-labels).
        
        DOM events have format:
        {
            Subject: "Meeting Title",
            Start: {DateTime: "2026-02-03T09:00:00", TimeZone: "UTC"},
            End: {DateTime: "2026-02-03T10:00:00", TimeZone: "UTC"},
            Organizer: {EmailAddress: {Name: "John Doe"}},
            IsRecurring: true/false,
            _rawLabel: "full aria-label text",
            _source: "dom"
        }
        """
        results = []
        logger.info(f"Processing {len(events)} DOM events, date range: {start} to {end}")

        for item in events:
            try:
                # Debug: log item structure
                source = item.get("_source")
                subject = item.get("Subject") or "(No Subject)"
                logger.debug(f"Processing event: {subject}, _source={source}, Start={item.get('Start')}")

                # Check if this is a DOM event
                if source != "dom":
                    logger.debug(f"Skipping non-DOM event: {subject} (source={source})")
                    continue

                # Parse dates
                start_obj = item.get("Start", {})
                end_obj = item.get("End", {})
                
                start_str = start_obj.get("DateTime", "") if isinstance(start_obj, dict) else ""
                end_str = end_obj.get("DateTime", "") if isinstance(end_obj, dict) else ""
                
                if not start_str:
                    # If no parsed date, try to parse from raw label
                    raw_label = item.get("_rawLabel", "")
                    logger.warning(f"DOM event missing date, raw label: {raw_label[:200]}")

                    # Try to parse date from raw label directly
                    parsed_date = self._parse_date_from_aria_label(raw_label)
                    if parsed_date:
                        start_str = parsed_date.get("start", "")
                        end_str = parsed_date.get("end", "")
                        logger.info(f"Parsed date from raw label: {start_str} to {end_str}")

                    if not start_str:
                        logger.debug(f"Skipping DOM event without date: {subject}")
                        continue
                
                # Parse datetime strings - DOM times are in LOCAL timezone, not UTC
                # Get local timezone for proper conversion
                import pytz
                from datetime import timezone as dt_timezone
                try:
                    # Try to get local timezone
                    local_tz = datetime.now().astimezone().tzinfo
                except Exception:
                    local_tz = dt_timezone.utc

                try:
                    if "T" in start_str:
                        # Check if already has timezone info
                        if "+" in start_str or "Z" in start_str:
                            start_dt = datetime.fromisoformat(start_str.replace("Z", "+00:00"))
                        else:
                            # DOM times are in LOCAL timezone, not UTC
                            start_dt = datetime.fromisoformat(start_str)
                            start_dt = start_dt.replace(tzinfo=local_tz)
                    else:
                        start_dt = datetime.fromisoformat(start_str + "T00:00:00")
                        start_dt = start_dt.replace(tzinfo=local_tz)
                    # Convert to UTC
                    start_dt = ensure_utc(start_dt)
                except ValueError:
                    logger.debug(f"Could not parse date {start_str} for event: {subject}")
                    continue

                try:
                    if end_str and "T" in end_str:
                        if "+" in end_str or "Z" in end_str:
                            end_dt = datetime.fromisoformat(end_str.replace("Z", "+00:00"))
                        else:
                            # DOM times are in LOCAL timezone
                            end_dt = datetime.fromisoformat(end_str)
                            end_dt = end_dt.replace(tzinfo=local_tz)
                    elif end_str:
                        end_dt = datetime.fromisoformat(end_str + "T23:59:00")
                        end_dt = end_dt.replace(tzinfo=local_tz)
                    else:
                        end_dt = start_dt + timedelta(hours=1)
                    # Convert to UTC
                    end_dt = ensure_utc(end_dt)
                except ValueError:
                    end_dt = start_dt + timedelta(hours=1)
                
                # Filter by date range
                if start_dt < start or start_dt > end:
                    logger.debug(f"Event {subject} outside date range: {start_dt} not in [{start}, {end}]")
                    continue

                # logger.info(f"✅ Event passed all filters: {subject} at {start_dt}")
                
                is_recurring = item.get("IsRecurring", False)
                
                # Parse organizer
                organizer = None
                org = item.get("Organizer", {})
                if org:
                    email_addr = org.get("EmailAddress", {}) if isinstance(org, dict) else {}
                    if email_addr:
                        organizer = Attendee(
                            email="",  # DOM doesn't give us email
                            name=email_addr.get("Name"),
                            is_organizer=True,
                        )
                
                # Generate a pseudo-ID from the event details
                import hashlib
                event_hash = hashlib.md5(
                    f"{subject}|{start_str}|{end_str}".encode()
                ).hexdigest()
                
                event = CalendarEvent(
                    id=f"dom-{event_hash}",
                    source_system="ews",
                    subject=subject,
                    start=ensure_utc(start_dt),
                    end=ensure_utc(end_dt),
                    is_all_day=False,  # Can't reliably determine from DOM
                    timezone="UTC",
                    organizer=organizer,
                    attendees=[],
                    location=None,  # Not available from DOM
                    status=EventStatus.CONFIRMED,
                    sensitivity="normal",
                    show_as="busy",
                    categories=[],
                    is_recurring=is_recurring,
                )
                
                if not self._should_skip(event):
                    results.append(event)
                    
            except Exception as e:
                logger.warning(f"Failed to parse DOM event: {e}")
                continue
        
        logger.info(f"Parsed {len(results)} events from DOM")
        return results

    def _parse_date_from_aria_label(self, label: str) -> dict | None:
        """Parse date/time from aria-label using various formats.

        Handles formats like:
        - "Subject, 09:00 to 10:00, Wednesday, February 05, 2026, ..."
        - "Subject, 09:00 bis 10:00, Mittwoch, 05. Februar 2026, ..."  (German)
        - "Subject, 09:00 à 10:00, mercredi 5 février 2026, ..."  (French)
        """
        if not label:
            return None

        # Parse time range - handle various formats
        time_patterns = [
            r'(\d{1,2}:\d{2})\s+to\s+(\d{1,2}:\d{2})',  # English
            r'(\d{1,2}:\d{2})\s+bis\s+(\d{1,2}:\d{2})',  # German
            r'(\d{1,2}:\d{2})\s+à\s+(\d{1,2}:\d{2})',   # French
            r'(\d{1,2}:\d{2})\s+-\s+(\d{1,2}:\d{2})',   # Dash separator
        ]

        start_time = None
        end_time = None
        for pattern in time_patterns:
            match = re.search(pattern, label)
            if match:
                start_time = match.group(1)
                end_time = match.group(2)
                break

        # Try to extract date - multiple approaches
        # Approach 1: Look for full date patterns
        date_patterns = [
            # English: "February 05, 2026" or "February 5, 2026"
            r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})',
            # German: "05. Februar 2026"
            r'(\d{1,2})\.\s*(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)\s+(\d{4})',
            # French: "5 février 2026"
            r'(\d{1,2})\s+(janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre)\s+(\d{4})',
            # ISO-like in label: "2026-02-05"
            r'(\d{4})-(\d{2})-(\d{2})',
            # Numeric: "05/02/2026" or "05.02.2026"
            r'(\d{1,2})[./](\d{1,2})[./](\d{4})',
        ]

        month_map = {
            # English
            'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
            'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
            # German
            'januar': 1, 'februar': 2, 'märz': 3, 'mai': 5, 'juni': 6,
            'juli': 7, 'august': 8, 'oktober': 10, 'dezember': 12,
            # French
            'janvier': 1, 'février': 2, 'mars': 3, 'avril': 4, 'juin': 6,
            'juillet': 7, 'août': 8, 'septembre': 9, 'octobre': 10, 'novembre': 11, 'décembre': 12,
        }

        year = None
        month = None
        day = None

        for i, pattern in enumerate(date_patterns):
            match = re.search(pattern, label, re.IGNORECASE)
            if match:
                if i == 0:  # English: Month Day, Year
                    month = month_map.get(match.group(1).lower())
                    day = int(match.group(2))
                    year = int(match.group(3))
                elif i == 1:  # German: Day. Month Year
                    day = int(match.group(1))
                    month = month_map.get(match.group(2).lower())
                    year = int(match.group(3))
                elif i == 2:  # French: Day Month Year
                    day = int(match.group(1))
                    month = month_map.get(match.group(2).lower())
                    year = int(match.group(3))
                elif i == 3:  # ISO: Year-Month-Day
                    year = int(match.group(1))
                    month = int(match.group(2))
                    day = int(match.group(3))
                elif i == 4:  # Numeric: Day/Month/Year (European)
                    day = int(match.group(1))
                    month = int(match.group(2))
                    year = int(match.group(3))
                break

        if not all([year, month, day]):
            logger.debug(f"Could not extract date from label: {label[:100]}...")
            return None

        # Build ISO date strings
        date_prefix = f"{year:04d}-{month:02d}-{day:02d}"
        start_str = f"{date_prefix}T{start_time or '00:00'}:00"
        end_str = f"{date_prefix}T{end_time or '23:59'}:00"

        return {"start": start_str, "end": end_str}

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

        try:
            data = resp.json()
        except json.JSONDecodeError:
            logger.warning(
                "OWA returned invalid JSON (likely expired cookies), "
                "deleting cookie cache and refreshing..."
            )
            self.selenium_auth.delete_cookie_cache()
            self._session = None
            self.selenium_auth.get_cookies(force_refresh=True)
            session = self._get_session()
            resp = session.post(url, json=payload, headers=action_headers, timeout=30)
            data = resp.json()
        resp_body = data.get("Body", {})
        messages = resp_body.get("ResponseMessages", {}).get("Items", [])

        if messages and messages[0].get("ResponseClass") != "Success":
            code = messages[0].get("ResponseCode", "Unknown")
            raise CalendarReadError(f"OWA {action} error: {code}")

        return data

    def list_calendars(self) -> list[Calendar]:
        """List calendars from Exchange."""
        # Initialize session to check canary status
        self._get_session()
        
        # For Office 365 without canary, just return a default calendar
        # since we'll use REST API for events anyway
        if self._is_office365() and not self._canary_token:
            logger.info("Using default calendar (no canary token for Office 365)")
            return [
                Calendar(
                    id="primary",
                    name="Calendar",
                    owner_email=self.config.primary_email,
                    source_system="ews",
                    is_default=True,
                    can_edit=True,
                )
            ]
        
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
            # Initialize session to check canary status
            self._get_session()
            
            start = ensure_utc(start_date) if start_date else ensure_utc(datetime(2020, 1, 1))
            end = ensure_utc(end_date) if end_date else ensure_utc(datetime(2030, 1, 1))

            # For Office 365 without canary, use browser-based API calls
            if self._is_office365() and not self._canary_token:
                logger.info("Using browser-based API calls for Office 365...")
                start_str = start.strftime("%Y-%m-%dT%H:%M:%SZ")
                end_str = end.strftime("%Y-%m-%dT%H:%M:%SZ")
                
                browser_events = self.selenium_auth.fetch_calendar_events_via_browser(
                    start_str, end_str
                )
                
                if browser_events is not None and len(browser_events) > 0:
                    # Check if these are DOM events or REST API events
                    first_event = browser_events[0] if browser_events else {}
                    if first_event.get("_source") == "dom":
                        return self._parse_dom_events(browser_events, start, end)
                    else:
                        return self._parse_rest_api_events(browser_events, start, end)
                
                logger.warning("Browser-based API failed, trying cookie-based approaches...")
                
                # Try REST API as fallback
                rest_events = self._rest_api_get_events(start, end)
                if rest_events is not None:
                    return self._parse_rest_api_events(rest_events, start, end)
                    
                raise CalendarReadError(
                    "Could not access Office 365 calendar. The OAuth tokens are not accessible. "
                    "Consider using the 'm365' account type with device code flow instead."
                )

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
