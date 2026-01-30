#!/usr/bin/env python3
"""Interactive setup helper for Calendar Sync configuration."""

import sys
from pathlib import Path


def main():
    print("\n" + "=" * 70)
    print("📅 Calendar Sync - Configuration Setup")
    print("=" * 70 + "\n")

    env_file = Path(".env")

    if env_file.exists():
        response = input("⚠️  .env file already exists. Overwrite? (y/N): ")
        if response.lower() != 'y':
            print("Setup cancelled.")
            return

    print("Let's configure your Microsoft 365 / Exchange settings.\n")
    print("You'll need to register an app in Azure Portal first:")
    print("https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps")
    print()

    # M365 Configuration
    print("─" * 70)
    print("Microsoft 365 Configuration")
    print("─" * 70)

    m365_tenant_id = input("\nM365 Tenant ID (from Azure Portal): ").strip()
    m365_client_id = input("M365 Client ID (from Azure Portal): ").strip()
    m365_client_secret = input("M365 Client Secret (optional, press Enter to skip): ").strip()

    if m365_tenant_id:
        m365_authority = f"https://login.microsoftonline.com/{m365_tenant_id}"
    else:
        m365_authority = input("M365 Authority URL: ").strip()

    # EWS Configuration
    print("\n" + "─" * 70)
    print("Exchange EWS Configuration")
    print("─" * 70)

    ews_server_url = input("\nEWS Server URL (e.g., https://mail.company.com/EWS/Exchange.asmx): ").strip()
    ews_client_id = input("EWS Client ID (from Azure Portal): ").strip()
    ews_tenant_id = input("EWS Tenant ID (from Azure Portal): ").strip()
    ews_primary_email = input("Your email address: ").strip()

    # Generate .env file
    env_content = f"""# Microsoft 365 Configuration
M365_TENANT_ID={m365_tenant_id}
M365_CLIENT_ID={m365_client_id}
M365_CLIENT_SECRET={m365_client_secret}
M365_AUTHORITY={m365_authority}
M365_SCOPES=Calendars.Read,Calendars.ReadWrite

# Exchange EWS Configuration
EWS_SERVER_URL={ews_server_url}
EWS_CLIENT_ID={ews_client_id}
EWS_TENANT_ID={ews_tenant_id}
EWS_PRIMARY_EMAIL={ews_primary_email}

# Token Cache Configuration
TOKEN_CACHE_PATH=.token_cache
TOKEN_CACHE_ENCRYPTED=true

# Logging Configuration
LOG_LEVEL=INFO
LOG_FILE=calendar_sync.log

# Sync Configuration
SYNC_DIRECTION=read_only
SYNC_LOOKBACK_DAYS=30
SYNC_LOOKAHEAD_DAYS=90
"""

    with open(".env", "w") as f:
        f.write(env_content)

    print("\n" + "=" * 70)
    print("✅ Configuration saved to .env")
    print("=" * 70)

    print("\n📋 Next steps:")
    print("1. Ensure your Azure AD app has these permissions:")
    print("   - Microsoft Graph: Calendars.Read, Calendars.ReadWrite")
    print("   - Exchange: EWS.AccessAsUser.All (if using EWS)")
    print("2. Set application type to 'Public client/native'")
    print("3. Run: uv run calendar-sync --source m365 --list-calendars")
    print()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nSetup cancelled.")
        sys.exit(0)
