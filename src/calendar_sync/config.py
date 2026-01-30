"""Configuration management for Calendar Sync application."""

from pathlib import Path
from typing import Any, Optional

import yaml
from dotenv import load_dotenv
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

load_dotenv()


class M365Config(BaseSettings):
    """Microsoft 365 configuration."""

    tenant_id: Optional[str] = Field(None, validation_alias="M365_TENANT_ID")
    client_id: Optional[str] = Field(None, validation_alias="M365_CLIENT_ID")
    client_secret: Optional[str] = Field(None, validation_alias="M365_CLIENT_SECRET")
    authority: Optional[str] = Field(None, validation_alias="M365_AUTHORITY")
    # Scopes are hardcoded - no need to configure
    scopes: list[str] = Field(default=["Calendars.Read", "Calendars.ReadWrite"])

    model_config = SettingsConfigDict(
        env_file=".env",
        extra="ignore",
        env_parse_none_str="",  # Treat empty string as None
    )


class EWSConfig(BaseSettings):
    """Exchange EWS configuration."""

    server_url: Optional[str] = Field(None, validation_alias="EWS_SERVER_URL")
    client_id: Optional[str] = Field(None, validation_alias="EWS_CLIENT_ID")
    tenant_id: Optional[str] = Field(None, validation_alias="EWS_TENANT_ID")
    primary_email: Optional[str] = Field(None, validation_alias="EWS_PRIMARY_EMAIL")

    # Authentication method: "oauth" or "selenium"
    auth_method: str = Field(default="oauth", validation_alias="EWS_AUTH_METHOD")

    # Selenium auth settings
    cookie_file: Path = Field(
        default=Path(".ews_cookies.json"), validation_alias="EWS_COOKIE_FILE"
    )

    model_config = SettingsConfigDict(
        env_file=".env",
        extra="ignore",
        env_parse_none_str="",  # Treat empty string as None
    )


class AppConfig(BaseSettings):
    """Application configuration."""

    # Auth configs
    m365: M365Config = Field(default_factory=M365Config)
    ews: EWSConfig = Field(default_factory=EWSConfig)

    # Token cache
    token_cache_path: Path = Field(
        default=Path(".token_cache"), validation_alias="TOKEN_CACHE_PATH"
    )
    token_cache_encrypted: bool = Field(
        default=True, validation_alias="TOKEN_CACHE_ENCRYPTED"
    )

    # Logging
    log_level: str = Field(default="INFO", validation_alias="LOG_LEVEL")
    log_file: Optional[Path] = Field(default=None, validation_alias="LOG_FILE")

    # Sync settings
    sync_direction: str = Field(default="read_only", validation_alias="SYNC_DIRECTION")
    sync_lookback_days: int = Field(default=30, validation_alias="SYNC_LOOKBACK_DAYS")
    sync_lookahead_days: int = Field(
        default=90, validation_alias="SYNC_LOOKAHEAD_DAYS"
    )

    model_config = SettingsConfigDict(
        env_file=".env",
        extra="ignore",
    )


class AccountConfig:
    """Configuration for a single calendar account."""

    def __init__(self, name: str, data: dict[str, Any]):
        self.name = name
        self.type: str = data.get("type", "m365")
        self.server_url: Optional[str] = data.get("server_url")
        self.primary_email: Optional[str] = data.get("primary_email")
        self.tenant_id: Optional[str] = data.get("tenant_id")
        self.client_id: Optional[str] = data.get("client_id")
        self.client_secret: Optional[str] = data.get("client_secret")
        self.cookie_file: Path = Path(data.get("cookie_file", f".ews_cookies_{name}.json"))
        self.required_cookies: list[str] = data.get("required_cookies", ["MRHSession"])
        self.prefix: str = data.get("prefix", "")
        self.category: Optional[str] = data.get("category")
        self.auth_method: str = data.get("auth_method", "selenium" if self.type == "ews_selenium" else "oauth")


class SyncConfig:
    """Multi-source sync configuration loaded from YAML."""

    def __init__(self, config_path: Path = Path("sync_config.yaml")):
        self.accounts: dict[str, AccountConfig] = {}
        self.sources: list[str] = []
        self.target: Optional[str] = None
        self.lookback_days: int = 0
        self.lookahead_days: int = 7
        self.skip_subjects: list[str] = []

        if config_path.exists():
            with open(config_path) as f:
                data = yaml.safe_load(f) or {}

            for name, acct_data in data.get("accounts", {}).items():
                self.accounts[name] = AccountConfig(name, acct_data)

            sync_data = data.get("sync", {})
            self.sources = sync_data.get("sources", [])
            self.target = sync_data.get("target")
            self.lookback_days = sync_data.get("lookback_days", 0)
            self.lookahead_days = sync_data.get("lookahead_days", 7)
            self.skip_subjects = [s.lower().strip() for s in data.get("skip_subjects", [])]

    @property
    def has_config(self) -> bool:
        return len(self.accounts) > 0


# Global config instances
config = AppConfig()
sync_config = SyncConfig()
