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

        self.config = config
        self.cache_manager = cache_manager
        self.client_secret = config.client_secret
        self.use_client_credentials = bool(config.client_id and config.client_secret)

        if self.use_client_credentials:
            # Client credentials flow (app-only) - uses application permissions
            logger.info("Initializing M365 auth with client credentials flow (app-only)")
            self.scopes = ["https://graph.microsoft.com/.default"]
            self.app = msal.ConfidentialClientApplication(
                client_id=client_id,
                client_credential=config.client_secret,
                authority=authority,
                token_cache=cache_manager.get_cache(),
            )
        else:
            # Device code flow (delegated) - uses delegated permissions
            logger.info(
                f"Initializing M365 auth with device code flow, client_id={'<default>' if not config.client_id else '<custom>'}"
            )
            self.scopes = [
                f"https://graph.microsoft.com/{scope}" for scope in config.scopes
            ]
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
        if self.use_client_credentials:
            # For client credentials, acquire_token_for_client handles caching internally
            # It will return a cached token if available and valid
            result = self.app.acquire_token_for_client(scopes=self.scopes)
            if result and "access_token" in result:
                # Only log if we didn't have to make a network call (token was cached)
                logger.debug("Token acquired for client credentials")
                return result["access_token"]
        else:
            # Delegated flow - check for user accounts
            accounts = self.app.get_accounts()
            if accounts:
                result = self.app.acquire_token_silent(
                    scopes=self.scopes,
                    account=accounts[0],
                )
                if result and "access_token" in result:
                    logger.debug("Token acquired from cache (delegated)")
                    return result["access_token"]
        return None

    def acquire_token_interactive(self) -> str:
        """
        Acquire token using appropriate flow based on configuration.

        - Client credentials flow when client_secret is configured
        - Device code flow otherwise (works in WSL/headless environments)

        Returns:
            Access token string

        Raises:
            AuthenticationError: If authentication fails
        """
        if self.use_client_credentials:
            return self._acquire_token_client_credentials()
        else:
            return self._acquire_token_device_code()

    def _acquire_token_client_credentials(self) -> str:
        """Acquire token using client credentials flow (app-only).

        MSAL automatically caches the token and returns it from cache if valid.
        """
        result = self.app.acquire_token_for_client(scopes=self.scopes)

        if "access_token" in result:
            # Only log at debug level since this is called frequently and MSAL handles caching
            logger.debug("Token acquired via client credentials flow")
            return result["access_token"]
        else:
            error_desc = result.get("error_description", "Unknown error")
            raise AuthenticationError(
                f"Client credentials authentication failed: {error_desc}"
            )

    def _acquire_token_device_code(self) -> str:
        """Acquire token using device code flow (delegated)."""
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
        if self.use_client_credentials:
            # For client credentials, acquire_token_for_client handles caching automatically
            return self._acquire_token_client_credentials()
        else:
            # For delegated flow, try silent first, then interactive
            token = self.acquire_token_silent()
            if token:
                return token
            return self.acquire_token_interactive()

    def clear_cache(self) -> None:
        """Clear cached tokens."""
        self.cache_manager.clear_cache()
