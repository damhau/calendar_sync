# Calendar Sync

Synchronize calendars between Exchange EWS (on-premise) and Microsoft 365.

## Features

- Read calendars from Exchange EWS (on-premise)
- Read calendars from Microsoft 365
- Write events to Microsoft 365
- Device code OAuth authentication (perfect for WSL/headless environments)
- Encrypted token caching
- Normalized event data model

## Requirements

- Python 3.11+
- Azure AD app registration (for M365 and EWS OAuth)
- Access to Exchange EWS server (for on-premise Exchange)

## Installation

### Using uv (recommended)

```bash
# Clone the repository
git clone <repository-url>
cd calendar_sync

# Install dependencies
uv sync

# Copy environment template
cp .env.template .env

# Edit .env with your configuration
```

### Using pip

```bash
pip install -e .
```

## Configuration

Create a `.env` file in the project root with the following configuration:

```env
# Microsoft 365 Configuration
M365_TENANT_ID=your-tenant-id-here
M365_CLIENT_ID=your-client-id-here
M365_CLIENT_SECRET=your-client-secret-here
M365_AUTHORITY=https://login.microsoftonline.com/your-tenant-id
M365_SCOPES=Calendars.Read,Calendars.ReadWrite

# Exchange EWS Configuration
EWS_SERVER_URL=https://your-exchange-server.com/EWS/Exchange.asmx
EWS_CLIENT_ID=your-ews-client-id-here
EWS_TENANT_ID=your-ews-tenant-id-here
EWS_PRIMARY_EMAIL=user@domain.com

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
```

## Azure AD App Setup

### For Microsoft 365

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to Azure Active Directory > App registrations
3. Click "New registration"
4. Set name (e.g., "Calendar Sync M365")
5. Set redirect URI: `http://localhost` (Public client/native)
6. Click "Register"
7. Note the "Application (client) ID" and "Directory (tenant) ID"
8. Go to "API permissions"
9. Add Microsoft Graph permissions:
   - `Calendars.Read`
   - `Calendars.ReadWrite`
10. Grant admin consent

### For Exchange EWS

Follow similar steps but use:
- API permissions: `Office 365 Exchange Online` > `EWS.AccessAsUser.All`

## Usage

### List Calendars

```bash
# List M365 calendars
uv run calendar-sync --source m365 --list-calendars

# List EWS calendars
uv run calendar-sync --source ews --list-calendars
```

### Preview Events

```bash
# Preview M365 events
uv run calendar-sync --source m365 --preview

# Preview EWS events
uv run calendar-sync --source ews --preview
```

### Sync Calendars

```bash
# Dry run (preview what would be synced)
uv run calendar-sync --source ews --target m365 --sync --dry-run

# Actual sync from EWS to M365
uv run calendar-sync --source ews --target m365 --sync

# Sync from M365 to M365 (different calendars)
uv run calendar-sync --source m365 --target m365 --sync
```

### Clear Token Cache

```bash
uv run calendar-sync --clear-cache
```

### Verbose Output

```bash
uv run calendar-sync --source m365 --list-calendars --verbose
```

## Project Structure

```
calendar_sync/
├── src/
│   └── calendar_sync/
│       ├── __init__.py
│       ├── __main__.py          # CLI entry point
│       ├── config.py            # Configuration management
│       ├── models/              # Data models
│       │   ├── event.py         # Normalized event model
│       │   └── calendar.py      # Calendar metadata
│       ├── auth/                # Authentication
│       │   ├── base.py          # Abstract auth interface
│       │   ├── msal_auth.py     # M365 authentication
│       │   ├── ews_auth.py      # EWS authentication
│       │   └── token_cache.py   # Token cache manager
│       ├── readers/             # Calendar readers
│       │   ├── base.py          # Abstract reader interface
│       │   ├── m365_reader.py   # M365 calendar reader
│       │   └── ews_reader.py    # EWS calendar reader
│       ├── writers/             # Calendar writers
│       │   ├── base.py          # Abstract writer interface
│       │   └── m365_writer.py   # M365 calendar writer
│       ├── sync/                # Sync engine
│       │   ├── engine.py        # Main sync orchestration
│       │   └── strategies.py    # Sync strategies
│       └── utils/               # Utilities
│           ├── logging.py       # Logging setup
│           ├── exceptions.py    # Custom exceptions
│           └── date_utils.py    # Date/time utilities
├── tests/                       # Tests
├── docs/                        # Documentation
├── .env.template                # Environment template
├── pyproject.toml              # Project configuration
└── README.md                    # This file
```

## Development

### Run Tests

```bash
uv run pytest
```

### Code Quality

```bash
# Lint and format
uv run ruff check .
uv run ruff format .

# Type checking
uv run mypy src/calendar_sync
```

## Troubleshooting

### Authentication Issues

- Ensure Azure AD app has correct permissions
- Check that redirect URI is set to `http://localhost`
- Clear token cache: `uv run calendar-sync --clear-cache`
- Check that tenant ID and client ID are correct

### EWS Connection Issues

- Verify EWS server URL is accessible
- Check that OAuth is enabled for your Exchange server
- Ensure email address matches authenticated user

### Token Cache Issues

On Linux, if LibSecret is not available:
```bash
# Install libsecret (Ubuntu/Debian)
sudo apt-get install libsecret-1-0 libsecret-1-dev
```

## License

MIT License

## Support

For issues and feature requests, please open an issue on GitHub.
