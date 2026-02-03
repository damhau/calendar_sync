#!/usr/bin/env python3
"""Test script to check if Calendars.Read works without admin consent for Helvetia."""

import msal
import requests

# Microsoft Graph PowerShell public client (well-known client ID)
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

# Try with "common" tenant (multi-tenant) and ONLY read permission
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["https://graph.microsoft.com/Calendars.Read"]  # Only read, not write

print("=" * 70)
print("Testing Helvetia M365 Access - Read-Only Permissions")
print("=" * 70)
print()
print("Client ID:", CLIENT_ID)
print("Authority:", AUTHORITY)
print("Scopes:", SCOPES)
print()

# Create MSAL app
app = msal.PublicClientApplication(
    client_id=CLIENT_ID,
    authority=AUTHORITY,
)

# Start device code flow
print("Starting device code flow...")
flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" not in flow:
    print("Failed to create device flow")
    exit(1)

# Display instructions
print()
print("=" * 70)
print("AUTHENTICATION REQUIRED")
print("=" * 70)
print(f"\n{flow['message']}\n")
print("=" * 70)
print()
print("IMPORTANT: Login with your Helvetia account (damien.hauser@helvetia.ch)")
print()

# Wait for authentication
result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    print("❌ Authentication failed!")
    print("Error:", result.get("error"))
    print("Description:", result.get("error_description"))
    exit(1)

print("✅ Authentication successful!")
print()

# Test API access
token = result["access_token"]
headers = {
    "Authorization": f"Bearer {token}",
    "Content-Type": "application/json",
}

print("Testing calendar access...")
response = requests.get(
    "https://graph.microsoft.com/v1.0/me/calendars",
    headers=headers,
)

if response.status_code == 200:
    calendars = response.json().get("value", [])
    print(f"✅ Success! Found {len(calendars)} calendar(s):")
    for cal in calendars:
        print(f"  - {cal.get('name')} (ID: {cal.get('id')})")
    print()
    print("🎉 Read-only access works without admin consent!")
else:
    print(f"❌ Failed to access calendars: {response.status_code}")
    print(response.text)
