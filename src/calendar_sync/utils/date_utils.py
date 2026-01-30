"""Date and time utilities for Calendar Sync application."""

from datetime import datetime, timedelta
from typing import Optional

import pytz


def ensure_utc(dt: datetime) -> datetime:
    """
    Ensure datetime is in UTC.

    Args:
        dt: Datetime to convert

    Returns:
        UTC datetime
    """
    if dt.tzinfo is None:
        return pytz.utc.localize(dt)
    return dt.astimezone(pytz.utc)


def get_sync_window(
    lookback_days: int = 30,
    lookahead_days: int = 90,
) -> tuple[datetime, datetime]:
    """
    Get the sync window (start, end) in UTC.

    Args:
        lookback_days: Days to look back from now
        lookahead_days: Days to look ahead from now

    Returns:
        Tuple of (start_date, end_date) in UTC
    """
    now = datetime.now(pytz.utc)
    start = now - timedelta(days=lookback_days)
    end = now + timedelta(days=lookahead_days)
    return start, end


def parse_recurrence_pattern(pattern_data: dict) -> Optional[dict]:
    """
    Parse recurrence pattern from various formats.

    Args:
        pattern_data: Recurrence pattern data from API

    Returns:
        Normalized recurrence pattern dict or None
    """
    # TODO: Implement when adding recurrence support
    return None
