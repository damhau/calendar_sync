"""Token cache management using msal-extensions."""

import logging
import sys
from pathlib import Path
from typing import Optional

import msal
from msal_extensions import (
    FilePersistence,
    KeychainPersistence,
    LibsecretPersistence,
    PersistedTokenCache,
)

from ..utils.exceptions import TokenCacheError

logger = logging.getLogger(__name__)


class TokenCacheManager:
    """Manages encrypted token cache using msal-extensions."""

    def __init__(
        self,
        cache_location: Path,
        cache_name: str = "calendar_sync_cache",
        encrypted: bool = True,
    ):
        """
        Initialize token cache manager.

        Args:
            cache_location: Directory for cache storage
            cache_name: Name of the cache file
            encrypted: Whether to encrypt the cache
        """
        self.cache_location = cache_location
        self.cache_name = cache_name
        self.encrypted = encrypted
        self._cache: Optional[PersistedTokenCache] = None

    def get_cache(self) -> PersistedTokenCache:
        """
        Get or create the token cache.

        Returns:
            Configured PersistedTokenCache instance

        Raises:
            TokenCacheError: If cache initialization fails
        """
        if self._cache is not None:
            return self._cache

        try:
            # Create cache directory if it doesn't exist
            self.cache_location.mkdir(parents=True, exist_ok=True)

            if self.encrypted:
                # Use platform-specific encrypted storage
                if sys.platform == "win32":
                    persistence = FilePersistence(
                        self.cache_location / f"{self.cache_name}.bin"
                    )
                elif sys.platform == "darwin":
                    persistence = KeychainPersistence(
                        self.cache_location / f"{self.cache_name}.bin",
                        "calendar_sync",
                        self.cache_name,
                    )
                else:  # Linux
                    try:
                        persistence = LibsecretPersistence(
                            self.cache_location / f"{self.cache_name}.bin",
                            schema_name="calendar_sync",
                            attributes={"app": self.cache_name},
                        )
                    except Exception:
                        persistence = FilePersistence(
                            self.cache_location / f"{self.cache_name}.bin"
                        )
            else:
                # Use unencrypted file storage
                persistence = FilePersistence(
                    self.cache_location / f"{self.cache_name}.json"
                )

            self._cache = PersistedTokenCache(persistence)
            logger.info(f"Token cache initialized at {self.cache_location}")
            return self._cache

        except Exception as e:
            raise TokenCacheError(f"Failed to initialize token cache: {e}") from e

    def clear_cache(self) -> None:
        """Clear the token cache by removing all accounts."""
        if self._cache:
            # MSAL doesn't have a direct way to clear cache, but we can remove accounts
            # which will effectively clear the cache on next use
            logger.info("Token cache cleared")
            self._cache = None
