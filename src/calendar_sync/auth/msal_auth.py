"""MSAL-based authentication for Microsoft 365."""

import logging
from typing import Optional

import msal

from ..config import M365Config
from ..utils.exceptions import AuthenticationError
from .base import AuthProvider
from .token_cache import TokenCacheManager

logger = logging.getLogger(__name__)


class M365AuthProvider(AuthProvider):
    """MSAL-based authentication for Microsoft 365."""

    def __init__(
        self,
        config: M365Config,
        cache_manager: TokenCacheManager,
    ):
        """
        Initialize M365 authentication provider.

        Args:
            config: Microsoft 365 configuration
            cache_manager: Token cache manager

        Raises:
            AuthenticationError: If required configuration is missing
        """
        # Use Microsoft Graph PowerShell public client ID if not specified
        # This is a well-known Microsoft client ID for personal/testing use
        client_id = config.client_id or "14d82eec-204b-4c2f-b7e8-296a70dab67e"

        # Default to common tenant for personal Microsoft accounts
        tenant_id = config.tenant_id or "common"
        authority = config.authority or f"https://login.microsoftonline.com/{tenant_id}"

        logger.info(
            f"Initializing M365 auth with client_id={'<default>' if not config.client_id else '<custom>'}"
        )

        self.config = config
        self.cache_manager = cache_manager
        self.scopes = [
            f"https://graph.microsoft.com/{scope}" for scope in config.scopes
        ]

        # Create MSAL PublicClientApplication
        self.app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
            token_cache=cache_manager.get_cache(),
        )

    def acquire_token_silent(self) -> Optional[str]:
        """
        Attempt to acquire token from cache.

        Returns:
            Access token if available in cache, None otherwise
        """
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(
                scopes=self.scopes,
                account=accounts[0],
            )
            if result and "access_token" in result:
                logger.info("Token acquired from cache")
                return result["access_token"]
        return None

    def acquire_token_interactive(self) -> str:
        """
        Acquire token using device code flow (works in WSL/headless environments).

        Returns:
            Access token string

        Raises:
            AuthenticationError: If authentication fails
        """
        logger.info("Starting device code authentication flow")

        # Use device code flow which works in WSL and headless environments
        flow = self.app.initiate_device_flow(scopes=self.scopes)

        if "user_code" not in flow:
            raise AuthenticationError(
                f"Failed to create device flow: {flow.get('error_description', 'Unknown error')}"
            )

        # Display the user code and URL to the user
        print("\n" + "=" * 70)
        print("AUTHENTICATION REQUIRED")
        print("=" * 70)
        print(f"\n{flow['message']}\n")
        print("=" * 70 + "\n")

        # Wait for the user to authenticate
        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            logger.info("Token acquired via device code flow")
            print("✓ Authentication successful!\n")
            return result["access_token"]
        else:
            error_desc = result.get("error_description", "Unknown error")
            raise AuthenticationError(
                f"Device code authentication failed: {error_desc}"
            )

    def get_access_token(self) -> str:
        """
        Get valid access token (cache first, then interactive).

        Returns:
            Valid access token string

        Raises:
            AuthenticationError: If authentication fails
        """
        token = self.acquire_token_silent()
        if token:
            return token
        return self.acquire_token_interactive()

    def clear_cache(self) -> None:
        """Clear cached tokens."""
        self.cache_manager.clear_cache()
