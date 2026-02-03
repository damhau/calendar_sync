"""Selenium-based authentication for Exchange behind F5/SAML/OIDC."""

import json
import logging
import time
from pathlib import Path
from typing import Optional

from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait

from ..utils.exceptions import AuthenticationError

logger = logging.getLogger(__name__)


class SeleniumEWSAuth:
    """
    Selenium-based authentication for Exchange Web Services behind F5 APM.

    This handles scenarios where Exchange is behind F5 load balancer with
    OIDC/SAML authentication that requires browser interaction.
    """

    def __init__(
        self,
        base_url: str,
        cookie_file: Path,
        required_cookies: Optional[list[str]] = None,
        browser: str = "chrome",
    ):
        """
        Initialize Selenium-based EWS authentication.

        Args:
            base_url: Base URL of the Exchange server (e.g., https://mail.ext.icrc.org)
            cookie_file: Path to store/load cookies
            required_cookies: List of required cookie names (e.g., ['MRHSession', 'FedAuth'])
            browser: Browser to use ('chrome' or 'edge')
        """
        self.base_url = base_url.rstrip("/")
        self.cookie_file = cookie_file
        self.required_cookies = required_cookies or ["MRHSession"]
        self.browser = browser.lower()
        self._cookies: Optional[dict[str, str]] = None

    def load_cookies(self) -> Optional[dict[str, str]]:
        """
        Load cookies from file.

        Returns:
            Dictionary of cookies if file exists and valid, None otherwise
        """
        if not self.cookie_file.exists():
            logger.info(f"Cookie file not found: {self.cookie_file}")
            return None

        try:
            with open(self.cookie_file, "r") as f:
                data = json.load(f)

            # Validate that required cookies are present
            if isinstance(data, dict):
                missing = [c for c in self.required_cookies if c not in data]
                if missing:
                    logger.warning(f"Missing required cookies: {missing}")
                    return None

                logger.info(f"Loaded {len(data)} cookies from {self.cookie_file}")
                return data

        except Exception as e:
            logger.warning(f"Failed to load cookies: {e}")

        return None

    def save_cookies(self, cookies: dict[str, str]) -> None:
        """
        Save cookies to file.

        Args:
            cookies: Dictionary of cookie name -> value
        """
        try:
            # Create parent directory if it doesn't exist
            self.cookie_file.parent.mkdir(parents=True, exist_ok=True)

            with open(self.cookie_file, "w") as f:
                json.dump(cookies, f, indent=2)

            logger.info(f"✅ Saved {len(cookies)} cookies to {self.cookie_file}")

        except Exception as e:
            logger.error(f"Failed to save cookies: {e}")
            raise AuthenticationError(f"Failed to save cookies: {e}") from e

    def fetch_cookies_from_browser(self) -> dict[str, str]:
        """
        Open browser and wait for user to login, then extract cookies.

        Returns:
            Dictionary of cookies

        Raises:
            AuthenticationError: If browser automation fails or cookies not found
        """
        browser_name = self.browser.capitalize()
        print(f"🌐 Opening {browser_name} to let you log in...")

        # Create browser-specific options and driver
        if self.browser == "edge":
            try:
                from selenium.webdriver.edge.options import Options
                options = Options()
                options.add_experimental_option("detach", False)
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                options.add_experimental_option("useAutomationExtension", False)
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                driver = webdriver.Edge(options=options)
            except ImportError:
                raise AuthenticationError(
                    "Edge WebDriver not available. Install with: pip install selenium[edge]"
                )
        else:  # chrome
            options = ChromeOptions()
            options.add_experimental_option("detach", False)
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--remote-debugging-port=9222")
            driver = webdriver.Chrome(options=options)
        # options = Options()
        # options.add_experimental_option("detach", False)
        # driver = webdriver.Chrome(options=options)
        # # driver.get(BASE_URL)
        try:
            login_url = f"{self.base_url}/owa"
            print(f"🔗 Navigating to {login_url}...")
            driver.get(login_url)

            # Wait for the page to actually load
            WebDriverWait(driver, 10).until(
                lambda d: d.current_url != "data:," and d.current_url != "about:blank"
            )

            print(f"✅ Loaded: {driver.current_url}")
            print()
            print("=" * 70)
            print("🔐 PLEASE COMPLETE AUTHENTICATION")
            print("=" * 70)
            print()
            print("1. Enter your username and password")
            print("2. Complete MFA (authenticator app, SMS, etc.)")
            print("3. Wait until you see your inbox/calendar")
            print()
            print(f"Waiting for authentication cookies: {', '.join(self.required_cookies)}")
            print("=" * 70)
            print()

            # Wait for all required cookies to appear AND for OWA to fully load
            cookies_found = {}
            owa_fully_loaded = False
            max_wait = 300  # 5 minutes timeout
            start_time = time.time()
            initial_wait_done = False

            print("⏳ Waiting for you to complete login and MFA...")
            print("   (Browser will close automatically once you reach your inbox)")
            print()

            while not owa_fully_loaded:
                if time.time() - start_time > max_wait:
                    raise AuthenticationError(
                        f"Timeout waiting for authentication. OWA loaded: {owa_fully_loaded}, URL: {driver.current_url[:50]}"
                    )

                current_url = driver.current_url

                # FIRST: Wait until we're completely off the login page
                if "login.microsoftonline" in current_url:
                    time.sleep(2)
                    continue

                # SECOND: Now we're on outlook.office.com - check for cookies
                if "outlook.office" in current_url:
                    # Check cookies only when we're on outlook domain
                    all_cookies = driver.get_cookies()
                    for cookie in all_cookies:
                        if cookie["name"] in self.required_cookies:
                            if cookie["name"] not in cookies_found:
                                cookies_found[cookie["name"]] = cookie["value"]
                                print(f"⏳ Found cookie on Outlook domain: {cookie['name']}")

                    # THIRD: Check if all cookies are present AND page is loaded
                    if len(cookies_found) >= len(self.required_cookies):
                        # Give extra time for page to fully load after redirect
                        if not initial_wait_done:
                            print("⏳ All cookies found on Outlook, waiting for page to fully load...")
                            initial_wait_done = True
                            time.sleep(5)  # Wait 5 seconds for page to settle

                        # Try to detect if the page is truly loaded by checking title
                        try:
                            page_title = driver.title.lower()
                            if any(keyword in page_title for keyword in ['outlook', 'inbox', 'mail', 'calendar']):
                                print(f"✅ OWA fully loaded (Title: {driver.title[:50]})")
                                owa_fully_loaded = True
                                break
                        except:
                            pass

                time.sleep(2)

            print(f"✅ Authentication complete! Closing browser...")

            # Save all cookies
            all_cookie_dict = {c["name"]: c["value"] for c in driver.get_cookies()}

            driver.quit()

            self.save_cookies(all_cookie_dict)
            return all_cookie_dict

        except Exception as e:
            logger.error(f"Browser automation failed: {e}")
            driver.quit()
            raise AuthenticationError(f"Failed to get cookies from browser: {e}") from e

    def get_cookies(self, force_refresh: bool = False) -> dict[str, str]:
        """
        Get valid cookies, using cache if available or browser if needed.

        Args:
            force_refresh: If True, force re-authentication even if cookies exist

        Returns:
            Dictionary of cookies

        Raises:
            AuthenticationError: If authentication fails
        """
        if force_refresh:
            logger.info("Forcing cookie refresh...")
            return self.fetch_cookies_from_browser()

        # Try to load from cache first
        cookies = self.load_cookies()

        if cookies:
            logger.info("Using cached cookies")
            return cookies

        # If no valid cookies, get new ones from browser
        logger.info("No valid cached cookies, launching browser...")
        return self.fetch_cookies_from_browser()

    def validate_cookies(self, cookies: dict[str, str]) -> bool:
        """
        Validate that cookies work by testing EWS endpoint.

        Args:
            cookies: Dictionary of cookies to test

        Returns:
            True if cookies are valid, False otherwise
        """
        import requests

        try:
            url = f"{self.base_url}/EWS/Exchange.asmx"
            response = requests.get(
                url,
                cookies=cookies,
                timeout=10,
                allow_redirects=False,
            )

            # If we get 200 or see WSDL content, cookies work
            if response.status_code == 200 or "wsdl" in response.text.lower():
                logger.info("✅ Cookies validated successfully")
                return True

            logger.warning(
                f"Cookie validation returned status {response.status_code}"
            )
            return False

        except Exception as e:
            logger.warning(f"Cookie validation failed: {e}")
            return False

    def clear_cookies(self) -> None:
        """Clear cached cookies."""
        if self.cookie_file.exists():
            self.cookie_file.unlink()
            logger.info("Cookie cache cleared")
