#!/usr/bin/env python3
"""Test different EWS authentication methods for ICRC mail server."""

import requests
from requests.auth import HTTPBasicAuth
import sys


def test_basic_auth():
    """Test if basic authentication works."""
    print("\n" + "=" * 70)
    print("Test 1: Basic Authentication")
    print("=" * 70)

    username = input("Enter your email (e.g., user@icrc.org): ")
    password = input("Enter your password: ")

    url = "https://mail.ext.icrc.org/EWS/Exchange.asmx"

    try:
        response = requests.get(
            url, auth=HTTPBasicAuth(username, password), timeout=10, verify=True
        )
        print(f"Status Code: {response.status_code}")
        print(f"Headers: {dict(response.headers)}")

        if response.status_code == 200:
            print("✅ Basic auth works!")
            return True
        elif response.status_code == 401:
            print("❌ Basic auth failed - credentials rejected or not supported")
            print(f"WWW-Authenticate: {response.headers.get('WWW-Authenticate', 'Not present')}")
        else:
            print(f"⚠️  Unexpected response: {response.status_code}")
            print(response.text[:500])

    except requests.exceptions.SSLError as e:
        print(f"❌ SSL Error: {e}")
    except requests.exceptions.ConnectionError as e:
        print(f"❌ Connection Error: {e}")
    except Exception as e:
        print(f"❌ Error: {e}")

    return False


def test_session_based():
    """Test if we can get session cookies from the login page."""
    print("\n" + "=" * 70)
    print("Test 2: Session-based Authentication (F5 APM)")
    print("=" * 70)

    print("This requires browser interaction...")
    print("F5 APM typically requires:")
    print("1. Login via web browser to get session cookies")
    print("2. Extract MRHSession or similar cookies")
    print("3. Use those cookies for EWS requests")
    print("\nWould you like to implement Selenium-based auth? (similar to your SharePoint code)")


def test_oauth_discovery():
    """Check if OAuth2 endpoints are available."""
    print("\n" + "=" * 70)
    print("Test 3: OAuth2/Modern Auth Discovery")
    print("=" * 70)

    # Try to find autodiscover
    urls_to_test = [
        "https://mail.ext.icrc.org/autodiscover/autodiscover.xml",
        "https://autodiscover.icrc.org/autodiscover/autodiscover.xml",
        "https://mail.ext.icrc.org/.well-known/openid-configuration",
    ]

    for url in urls_to_test:
        print(f"\nTrying: {url}")
        try:
            response = requests.get(url, timeout=5, verify=True)
            print(f"  Status: {response.status_code}")
            if response.status_code == 200:
                print(f"  ✅ Found! Content-Type: {response.headers.get('Content-Type')}")
                print(f"  Response preview: {response.text[:200]}")
        except Exception as e:
            print(f"  ❌ {type(e).__name__}: {e}")


def check_ews_endpoint():
    """Check what the EWS endpoint returns without auth."""
    print("\n" + "=" * 70)
    print("Test 4: EWS Endpoint Check (no auth)")
    print("=" * 70)

    url = "https://mail.ext.icrc.org/EWS/Exchange.asmx"

    try:
        response = requests.get(url, timeout=10, verify=True, allow_redirects=True)
        print(f"Status Code: {response.status_code}")
        print(f"Final URL: {response.url}")
        print(f"Headers: {dict(response.headers)}")

        if "Location" in response.headers:
            print(f"\n⚠️  Redirected to: {response.headers['Location']}")
            print("This suggests F5 is intercepting and redirecting to login")

        print(f"\nResponse preview:\n{response.text[:500]}")

    except Exception as e:
        print(f"❌ Error: {e}")


def main():
    print("\n" + "=" * 70)
    print("EWS Authentication Testing for ICRC Mail Server")
    print("=" * 70)

    # Test endpoint first
    check_ews_endpoint()

    # Test OAuth discovery
    test_oauth_discovery()

    # Test basic auth if user wants
    response = input("\n\nTry basic authentication? (y/N): ")
    if response.lower() == "y":
        test_basic_auth()

    # Info about session-based auth
    test_session_based()

    print("\n" + "=" * 70)
    print("Summary:")
    print("=" * 70)
    print("Based on the F5 APM setup, you likely need one of:")
    print("1. Selenium-based authentication (like your SharePoint code)")
    print("2. Manual session cookie extraction")
    print("3. Basic auth (if enabled by IT)")
    print("\nNext steps: Let me know which approach you'd like to implement!")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nTest cancelled.")
        sys.exit(0)
