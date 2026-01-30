"""Abstract base class for authentication providers."""

from abc import ABC, abstractmethod
from typing import Optional


class AuthProvider(ABC):
    """Abstract base class for authentication providers."""

    @abstractmethod
    def acquire_token_interactive(self) -> str:
        """
        Acquire token using device code flow (works in WSL/headless).

        Returns:
            Access token string

        Raises:
            AuthenticationError: If authentication fails
        """

    @abstractmethod
    def acquire_token_silent(self) -> Optional[str]:
        """
        Attempt to acquire token silently from cache.

        Returns:
            Access token string if available in cache, None otherwise
        """

    @abstractmethod
    def get_access_token(self) -> str:
        """
        Get a valid access token (silent first, then interactive).

        Returns:
            Valid access token string

        Raises:
            AuthenticationError: If authentication fails
        """

    @abstractmethod
    def clear_cache(self) -> None:
        """Clear cached tokens."""
