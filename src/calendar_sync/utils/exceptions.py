"""Custom exceptions for Calendar Sync application."""


class CalendarSyncError(Exception):
    """Base exception for calendar sync errors."""


class AuthenticationError(CalendarSyncError):
    """Raised when authentication fails."""


class CalendarReadError(CalendarSyncError):
    """Raised when reading calendar fails."""


class CalendarWriteError(CalendarSyncError):
    """Raised when writing calendar fails."""


class TokenCacheError(CalendarSyncError):
    """Raised when token cache operations fail."""


class ConfigurationError(CalendarSyncError):
    """Raised when configuration is invalid."""
