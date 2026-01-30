"""OAuth authentication for Exchange EWS using MSAL."""

import logging
from typing import Optional

import msal
from exchangelib import OAuth2Credentials

from ..config import EWSConfig
from ..utils.exceptions import AuthenticationError
from .base import AuthProvider
from .token_cache import TokenCacheManager

logger = logging.getLogger(__name__)


class EWSAuthProvider(AuthProvider):
    """OAuth authentication for Exchange EWS using MSAL."""

    def __init__(
        self,
        config: EWSConfig,
        cache_manager: TokenCacheManager,
    ):
        """
        Initialize EWS authentication provider.

        Args:
            config: Exchange EWS configuration
            cache_manager: Token cache manager

        Raises:
            AuthenticationError: If required configuration is missing
        """
        # Use Microsoft Graph PowerShell public client ID if not specified
        client_id = config.client_id or "14d82eec-204b-4c2f-b7e8-296a70dab67e"
        tenant_id = config.tenant_id or "common"

        # Validate required EWS-specific configuration
        if not config.server_url:
            raise AuthenticationError(
                "EWS_SERVER_URL is required. Please set it in your .env file.\n"
                "For Microsoft 365: https://outlook.office365.com/EWS/Exchange.asmx\n"
                "For on-premise: https://your-exchange-server.com/EWS/Exchange.asmx"
            )
        if not config.primary_email:
            raise AuthenticationError(
                "EWS_PRIMARY_EMAIL is required. Please set it in your .env file.\n"
                "Format: your-email@domain.com"
            )

        logger.info(
            f"Initializing EWS auth with client_id={'<default>' if not config.client_id else '<custom>'}"
        )

        self.config = config
        self.cache_manager = cache_manager
        self.scopes = ["https://outlook.office365.com/EWS.AccessAsUser.All"]

        # Create MSAL PublicClientApplication
        self.app = msal.PublicClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            token_cache=cache_manager.get_cache(),
        )
        self._credentials: Optional[OAuth2Credentials] = None

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
                logger.info("EWS token acquired from cache")
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
        logger.info("Starting EWS device code authentication")

        # Use device code flow which works in WSL and headless environments
        flow = self.app.initiate_device_flow(scopes=self.scopes)

        if "user_code" not in flow:
            raise AuthenticationError(
                f"Failed to create device flow: {flow.get('error_description', 'Unknown error')}"
            )

        # Display the user code and URL to the user
        print("\n" + "=" * 70)
        print("EWS AUTHENTICATION REQUIRED")
        print("=" * 70)
        print(f"\n{flow['message']}\n")
        print("=" * 70 + "\n")

        # Wait for the user to authenticate
        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            logger.info("EWS token acquired via device code flow")
            print("✓ Authentication successful!\n")
            return result["access_token"]
        else:
            error_desc = result.get("error_description", "Unknown error")
            raise AuthenticationError(f"EWS authentication failed: {error_desc}")

    def get_access_token(self) -> str:
        """
        Get valid access token.

        Returns:
            Valid access token string

        Raises:
            AuthenticationError: If authentication fails
        """
        token = self.acquire_token_silent()
        if token:
            return token
        return self.acquire_token_interactive()

    def get_credentials(self) -> OAuth2Credentials:
        """
        Get OAuth2Credentials for exchangelib.

        Returns:
            OAuth2Credentials instance configured for EWS

        Raises:
            AuthenticationError: If token acquisition fails
        """
        access_token = self.get_access_token()
        return OAuth2Credentials(
            client_id=self.config.client_id,
            client_secret=None,  # Not needed for public client flow
            tenant_id=self.config.tenant_id,
            identity=self.config.primary_email,
        )

    def clear_cache(self) -> None:
        """Clear cached tokens."""
        self.cache_manager.clear_cache()
