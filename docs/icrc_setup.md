# ICRC Exchange Server Setup Guide

This guide explains how to use the calendar sync application with ICRC's Exchange server, which is protected by F5 APM with OIDC/SAML authentication.

## Overview

ICRC's Exchange server (`https://mail.ext.icrc.org`) uses F5 Access Policy Manager (APM) with OIDC/SAML authentication. Standard OAuth2 authentication won't work because F5 intercepts all requests and requires browser-based authentication.

**Solution:** Selenium-based authentication that:
1. Opens Chrome browser
2. Lets you login through F5/OIDC/SAML
3. Extracts and caches session cookies
4. Uses those cookies for EWS API requests

## Prerequisites

1. Chrome browser installed
2. Your ICRC credentials
3. Network access to `mail.ext.icrc.org`

## Configuration

### Step 1: Update Your `.env` File

Replace the EWS configuration in your [.env](../.env) file with:

```env
# ICRC Exchange Configuration
EWS_SERVER_URL=https://mail.ext.icrc.org
EWS_PRIMARY_EMAIL=your.name@icrc.org
EWS_AUTH_METHOD=selenium
EWS_COOKIE_FILE=.ews_cookies.json

# These are not needed for Selenium auth
EWS_CLIENT_ID=
EWS_TENANT_ID=
```

**Important:** Replace `your.name@icrc.org` with your actual ICRC email address.

### Step 2: First Run - Authentication

The first time you run the app with ICRC, it will:

1. Open Chrome browser
2. Navigate to `https://mail.ext.icrc.org/owa`
3. Wait for you to complete the F5/OIDC/SAML login
4. Extract and save session cookies to `.ews_cookies.json`

Run:

```bash
uv run calendar-sync --source ews --list-calendars
```

You'll see:

```
🌐 Opening Chrome for authentication...
======================================================================
🔐 AUTHENTICATION REQUIRED
======================================================================

Please login in the opened Chrome window.
URL: https://mail.ext.icrc.org/owa

Waiting for required cookies...
Required: MRHSession
======================================================================
```

**What to do:**
1. Complete the login in the Chrome window
2. Wait until you reach your inbox/OWA page
3. The app will automatically detect the session cookies
4. Chrome will close automatically
5. Cookies are saved to `.ews_cookies.json`

### Step 3: Subsequent Runs

After the first authentication, cookies are cached. Future runs will:
1. Load cookies from `.ews_cookies.json`
2. Validate they still work
3. Use them for EWS requests
4. **No browser opening needed!**

If cookies expire (usually after a few hours), the browser will automatically open again.

## Usage Examples

### List Your ICRC Calendars

```bash
uv run calendar-sync --source ews --list-calendars
```

### Preview ICRC Calendar Events

```bash
uv run calendar-sync --source ews --preview
```

This shows events from the last 30 days and next 90 days.

### Sync ICRC Calendar to Microsoft 365

```bash
# Dry run first
uv run calendar-sync --source ews --target m365 --sync --dry-run

# Actually sync
uv run calendar-sync --source ews --target m365 --sync
```

This copies your ICRC calendar events to your personal Microsoft 365 calendar.

## Troubleshooting

### Cookies Expired

**Symptom:** Getting authentication errors or redirects

**Solution:** Clear cookies and re-authenticate:

```bash
# Delete cookie file
rm .ews_cookies.json

# Run again - will open browser
uv run calendar-sync --source ews --list-calendars
```

### Chrome Not Found

**Symptom:** `ChromeDriver` or Chrome browser errors

**Solution:** Install Chrome:

```bash
# Ubuntu/Debian
sudo apt update
sudo apt install google-chrome-stable

# Or download from: https://www.google.com/chrome/
```

### F5 Session Timeout

**Symptom:** Cookies work initially but fail after a few hours

**Explanation:** F5 APM sessions expire. The app will automatically:
1. Detect invalid cookies
2. Open browser for re-authentication
3. Save new cookies

### Connection Issues

**Symptom:** `Connection reset` or `SSL errors`

**Possible causes:**
1. Not on ICRC network/VPN
2. F5 blocking non-browser user agents
3. Network firewall

**Solution:**
- Ensure you're on ICRC network or VPN
- Check with IT if you continue having issues

## How It Works

### Authentication Flow

```
1. App checks .ews_cookies.json
   ├─ Exists & valid? → Use cookies
   └─ Missing/invalid? → Open browser

2. Browser opens to mail.ext.icrc.org/owa
   ├─ F5 redirects to OIDC/SAML login
   ├─ You complete authentication
   └─ F5 sets MRHSession cookie

3. App detects cookie
   ├─ Saves to .ews_cookies.json
   └─ Closes browser

4. Future requests use cached cookies
   └─ Added to all EWS API calls
```

### Session Management

- **Cookies stored:** `.ews_cookies.json` (excluded from git)
- **Primary cookie:** `MRHSession` (F5 APM session ID)
- **Typical lifetime:** A few hours (set by F5 policy)
- **Auto-refresh:** Yes, browser opens when needed

## Security Notes

1. **Cookie Storage:** Cookies are stored in plain JSON
   - Keep `.ews_cookies.json` secure
   - Already in `.gitignore` - won't be committed
   - Delete when not needed: `rm .ews_cookies.json`

2. **Session Scope:** Cookies only work for:
   - The machine they were created on
   - While F5 session is active
   - With the same network access

3. **Best Practices:**
   - Don't share `.ews_cookies.json`
   - Delete cookies when finished: `rm .ews_cookies.json`
   - Re-authenticate if suspicious activity

## Comparison: ICRC vs. M365

| Feature | ICRC Exchange | Microsoft 365 |
|---------|---------------|---------------|
| **Auth Method** | Selenium (F5 APM) | Device Code (MSAL) |
| **Browser Needed** | Yes (first time) | No (code in terminal) |
| **Session Duration** | Few hours | Weeks/months |
| **Offline Support** | While session active | Yes (token refresh) |
| **Setup Complexity** | Medium | Easy |

## Advanced: Headless Mode

For automated/scheduled syncs, you can use headless Chrome:

1. Manually authenticate once
2. Save cookies
3. Run in headless mode (coming soon)

This allows scheduled syncs without browser interaction.

## Support

For issues specific to ICRC Exchange access:
- Contact ICRC IT support
- Check F5 APM logs if you have access
- Verify VPN/network connectivity

For app issues:
- Open an issue on GitHub
- Check logs in `calendar_sync.log`
- Run with `--verbose` flag

## Example: Full Sync Workflow

```bash
# 1. Configure for ICRC
nano .env  # Set EWS_SERVER_URL, EWS_PRIMARY_EMAIL, EWS_AUTH_METHOD=selenium

# 2. First authentication
uv run calendar-sync --source ews --list-calendars
# → Opens browser, you login, cookies saved

# 3. Preview events
uv run calendar-sync --source ews --preview
# → Uses cached cookies, no browser

# 4. Sync to M365 (dry run)
uv run calendar-sync --source ews --target m365 --sync --dry-run
# → Shows what would be synced

# 5. Actually sync
uv run calendar-sync --source ews --target m365 --sync
# → Copies events to M365

# 6. Later runs use cached cookies
uv run calendar-sync --source ews --preview
# → Instant, no browser needed
```

That's it! You now have automated calendar sync between ICRC Exchange and Microsoft 365. 🎉
