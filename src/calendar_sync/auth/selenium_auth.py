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
        use_browser_api: bool = False,
        headless: bool = False,
    ):
        """
        Initialize Selenium-based EWS authentication.

        Args:
            base_url: Base URL of the Exchange server (e.g., https://mail.ext.icrc.org)
            cookie_file: Path to store/load cookies
            required_cookies: List of required cookie names (e.g., ['MRHSession', 'FedAuth'])
            browser: Browser to use ('chrome' or 'edge')
            use_browser_api: If True, keep browser open for API calls instead of using cookies
            headless: If True, run browser in headless mode (no visible UI)
        """
        self.base_url = base_url.rstrip("/")
        self.cookie_file = cookie_file
        self.required_cookies = required_cookies or ["MRHSession"]
        self.browser = browser.lower()
        self.use_browser_api = use_browser_api
        self.headless = headless
        self._cookies: Optional[dict[str, str]] = None
        self._driver: Optional[webdriver.Chrome] = None  # Keep browser instance

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
        headless_msg = " (headless)" if self.headless else ""
        print(f"🌐 Opening {browser_name}{headless_msg} to let you log in...")

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
                if self.headless:
                    options.add_argument("--headless=new")
                    options.add_argument("--window-size=1920,1080")
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
            if self.headless:
                options.add_argument("--headless=new")
                options.add_argument("--window-size=1920,1080")
            driver = webdriver.Chrome(options=options)
        # options = Options()
        # options.add_experimental_option("detach", False)
        # driver = webdriver.Chrome(options=options)
        # # driver.get(BASE_URL)
        try:
            login_url = f"{self.base_url}/owa/?path=/calendar"
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

                # SECOND: Now we're on the target domain - check for cookies
                # This works for both Office 365 (outlook.office.com) and on-premise (e.g., mail.ext.icrc.org)
                base_domain = self.base_url.replace("https://", "").replace("http://", "").split("/")[0]
                is_on_target_domain = base_domain in current_url or "outlook.office" in current_url
                
                if is_on_target_domain:
                    # Check cookies only when we're on the target domain
                    all_cookies = driver.get_cookies()
                    for cookie in all_cookies:
                        if cookie["name"] in self.required_cookies:
                            if cookie["name"] not in cookies_found:
                                cookies_found[cookie["name"]] = cookie["value"]
                                print(f"⏳ Found cookie: {cookie['name']}")

                    # THIRD: Check if all cookies are present AND page is loaded
                    if len(cookies_found) >= len(self.required_cookies):
                        # Give extra time for page to fully load after redirect
                        if not initial_wait_done:
                            print("⏳ All cookies found, waiting for page to fully load...")
                            initial_wait_done = True
                            time.sleep(5)  # Wait 5 seconds for page to settle

                        # Try to detect if the page is truly loaded by checking title
                        try:
                            page_title = driver.title.lower()
                            if any(keyword in page_title for keyword in ['outlook', 'inbox', 'mail', 'calendar', 'owa']):
                                print(f"✅ OWA fully loaded (Title: {driver.title[:50]})")
                                owa_fully_loaded = True
                                break
                        except:
                            pass

                time.sleep(2)

            print(f"✅ Authentication complete! Extracting tokens...")

            # Save all cookies
            all_cookie_dict = {c["name"]: c["value"] for c in driver.get_cookies()}

            # CRITICAL: Extract X-OWA-CANARY token - required for all OWA API calls
            # For Office 365 OWA, the canary is obtained via fetch API with credentials
            canary_token = None
            
            # Method 1: Try fetch to get canary from response headers
            try:
                print("⏳ Fetching X-OWA-CANARY via service endpoint...")
                canary_token = driver.execute_script("""
                    return new Promise((resolve) => {
                        fetch('/owa/service.svc?action=GetOwaUserConfiguration', {
                            method: 'POST',
                            credentials: 'include',
                            headers: {
                                'Content-Type': 'application/json',
                                'Action': 'GetOwaUserConfiguration'
                            },
                            body: JSON.stringify({
                                "__type": "GetOwaUserConfigurationRequest:#Exchange",
                                "Header": {
                                    "__type": "JsonRequestHeaders:#Exchange",
                                    "RequestServerVersion": "Exchange2013"
                                },
                                "Body": {
                                    "__type": "GetOwaUserConfigurationRequest:#Exchange",
                                    "UserConfigurationName": {"__type": "UserConfigurationName:#Exchange", "Name": "OWA.ViewStateConfiguration", "DistinguishedFolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": "root"}}
                                }
                            })
                        })
                        .then(resp => {
                            // The canary is in the response header
                            let canary = resp.headers.get('X-OWA-CANARY');
                            if (canary) {
                                resolve(canary);
                            } else {
                                // Also check cookies that might have been set
                                let cookies = document.cookie.split(';');
                                for (let c of cookies) {
                                    if (c.trim().startsWith('X-OWA-CANARY=')) {
                                        resolve(c.trim().split('=')[1]);
                                        return;
                                    }
                                }
                                resolve(null);
                            }
                        })
                        .catch(() => resolve(null));
                        
                        // Timeout after 5 seconds
                        setTimeout(() => resolve(null), 5000);
                    });
                """)
                if canary_token:
                    print(f"✅ Found X-OWA-CANARY from service endpoint")
            except Exception as e:
                logger.debug(f"Could not get canary from fetch: {e}")

            # Method 2: Check if it appeared as a cookie after the fetch
            if not canary_token:
                time.sleep(1)
                for cookie in driver.get_cookies():
                    if cookie["name"] == "X-OWA-CANARY":
                        canary_token = cookie["value"]
                        all_cookie_dict["X-OWA-CANARY"] = canary_token
                        print(f"✅ Found X-OWA-CANARY in cookies after fetch")
                        break

            # Method 3: Try localStorage
            if not canary_token:
                try:
                    canary_token = driver.execute_script(
                        "return window.localStorage.getItem('x-owa-canary') || "
                        "window.localStorage.getItem('X-OWA-CANARY');"
                    )
                    if canary_token:
                        print(f"✅ Found X-OWA-CANARY in localStorage")
                except Exception as e:
                    logger.debug(f"Could not get canary from localStorage: {e}")

            # Method 4: Try to extract from OWA's JavaScript boot data
            if not canary_token:
                try:
                    canary_token = driver.execute_script("""
                        try {
                            // Try multiple known locations where OWA stores the canary
                            if (window.g_CanaryValue) return window.g_CanaryValue;
                            if (window.odataCanary) return window.odataCanary;
                            if (window.__owa_boot && window.__owa_boot.canary) return window.__owa_boot.canary;
                            if (typeof Boot !== 'undefined' && Boot.canary) return Boot.canary;
                            
                            // Search in script tags for canary patterns
                            var scripts = document.getElementsByTagName('script');
                            for (var i = 0; i < scripts.length; i++) {
                                var content = scripts[i].innerHTML;
                                var match = content.match(/"canary"\\s*:\\s*"([^"]+)"/);
                                if (match) return match[1];
                                match = content.match(/CanaryValue\\s*=\\s*"([^"]+)"/);
                                if (match) return match[1];
                                match = content.match(/x-owa-canary['"\\s:]+['"]([^'"]+)['"]/i);
                                if (match) return match[1];
                            }
                            return null;
                        } catch(e) { return null; }
                    """)
                    if canary_token:
                        print(f"✅ Found X-OWA-CANARY in page JavaScript")
                except Exception as e:
                    logger.debug(f"Could not get canary from JS context: {e}")

            # Method 5: Try sessionStorage
            if not canary_token:
                try:
                    canary_token = driver.execute_script("""
                        // Check sessionStorage
                        for (let i = 0; i < sessionStorage.length; i++) {
                            let key = sessionStorage.key(i);
                            if (key.toLowerCase().includes('canary')) {
                                return sessionStorage.getItem(key);
                            }
                        }
                        // Check for OWA boot data in various forms
                        try {
                            if (window.O365Shell && window.O365Shell.FlexPane) {
                                let data = window.O365Shell.FlexPane.HeaderButton;
                                if (data && data.canary) return data.canary;
                            }
                        } catch(e) {}
                        return null;
                    """)
                    if canary_token:
                        print(f"✅ Found X-OWA-CANARY in sessionStorage")
                except Exception as e:
                    logger.debug(f"Could not get canary from sessionStorage: {e}")

            # Method 6: Navigate to calendar and capture canary from network
            if not canary_token:
                try:
                    print("⏳ Navigating to calendar to trigger canary generation...")
                    driver.get(f"{self.base_url}/owa/?path=/calendar")
                    time.sleep(3)
                    
                    # Check cookies again
                    for cookie in driver.get_cookies():
                        if cookie["name"] == "X-OWA-CANARY":
                            canary_token = cookie["value"]
                            all_cookie_dict["X-OWA-CANARY"] = canary_token
                            print(f"✅ Found X-OWA-CANARY after calendar navigation")
                            break
                    
                    # If still no canary, try extracting from OWA's internal state
                    if not canary_token:
                        canary_token = driver.execute_script("""
                            // Try to find canary in OWA's React state or internal objects
                            try {
                                // Check window.__PRELOADED_STATE__ (common in React apps)
                                if (window.__PRELOADED_STATE__ && window.__PRELOADED_STATE__.session) {
                                    return window.__PRELOADED_STATE__.session.canary;
                                }
                            } catch(e) {}
                            
                            // Try to get it from a network request by triggering one
                            try {
                                // Look for any element that might have the canary as a data attribute
                                let allElements = document.querySelectorAll('[data-canary]');
                                if (allElements.length > 0) {
                                    return allElements[0].getAttribute('data-canary');
                                }
                            } catch(e) {}
                            
                            // Check all script tags for canary patterns
                            let scripts = document.getElementsByTagName('script');
                            for (let i = 0; i < scripts.length; i++) {
                                let content = scripts[i].textContent || scripts[i].innerHTML;
                                if (content) {
                                    // Look for canary in JSON-like structures
                                    let patterns = [
                                        /"canary"\\s*:\\s*"([^"]+)"/,
                                        /'canary'\\s*:\\s*'([^']+)'/,
                                        /canary['"]\\s*:\\s*['"]([\w\\-\\.]+)['"]/,
                                        /X-OWA-CANARY['"]\\s*:\\s*['"]([\w\\-\\.]+)['"]/i
                                    ];
                                    for (let pattern of patterns) {
                                        let match = content.match(pattern);
                                        if (match) return match[1];
                                    }
                                }
                            }
                            return null;
                        """)
                        if canary_token:
                            print(f"✅ Found X-OWA-CANARY from page analysis")
                    
                    # Update all cookies after navigation
                    if canary_token:
                        all_cookie_dict = {c["name"]: c["value"] for c in driver.get_cookies()}
                except Exception as e:
                    logger.debug(f"Calendar navigation failed: {e}")
                except Exception as e:
                    logger.debug(f"Calendar navigation failed: {e}")

            if canary_token:
                all_cookie_dict["X-OWA-CANARY"] = canary_token
                print(f"✅ X-OWA-CANARY token captured: {canary_token[:30]}...")
            else:
                print("⚠️  WARNING: Could not find X-OWA-CANARY token!")
                if self.use_browser_api:
                    print("   Browser will stay open for API calls.")
                else:
                    print("   OWA API calls may fail with 401 Unauthorized.")

            # If use_browser_api is enabled, keep the browser open
            if self.use_browser_api and not canary_token:
                print("✅ Keeping browser open for API calls (use_browser_api=true)")
                self._driver = driver
                # Navigate to calendar for API calls
                driver.get(f"{self.base_url}/owa/?path=/calendar")
                time.sleep(2)
            else:
                print(f"✅ Closing browser...")
                driver.quit()

            self.save_cookies(all_cookie_dict)
            return all_cookie_dict

        except Exception as e:
            logger.error(f"Browser automation failed: {e}")
            try:
                driver.quit()
            except:
                pass
            raise AuthenticationError(f"Failed to get cookies from browser: {e}") from e
    
    def has_browser(self) -> bool:
        """Check if browser is available for API calls."""
        return self._driver is not None
    
    def close_browser(self) -> None:
        """Close the browser if it's open."""
        if self._driver:
            try:
                self._driver.quit()
            except:
                pass
            self._driver = None

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
            
            # For use_browser_api mode, we need to launch browser for API calls
            # even when using cached cookies (Office 365 requires browser context)
            if self.use_browser_api and "X-OWA-CANARY" not in cookies:
                logger.info("Launching browser for API calls (use_browser_api mode)...")
                self._launch_browser_for_api(cookies)
            
            return cookies

        # If no valid cookies, get new ones from browser
        logger.info("No valid cached cookies, launching browser...")
        return self.fetch_cookies_from_browser()
    
    def _launch_browser_for_api(self, cookies: dict[str, str]) -> None:
        """
        Launch browser with existing cookies for making API calls.

        This is used when use_browser_api is enabled and we have cached cookies
        but no canary token (Office 365 scenario).
        """
        from selenium.webdriver.chrome.options import Options as ChromeOptions

        headless_msg = " (headless)" if self.headless else ""
        print(f"🌐 Launching browser{headless_msg} for API calls...")

        if self.browser == "edge":
            try:
                from selenium.webdriver.edge.options import Options
                options = Options()
                options.add_experimental_option("detach", False)
                options.add_argument("--disable-blink-features=AutomationControlled")
                if self.headless:
                    options.add_argument("--headless=new")
                    options.add_argument("--window-size=1920,1080")
                driver = webdriver.Edge(options=options)
            except ImportError:
                raise AuthenticationError("Edge WebDriver not available")
        else:
            options = ChromeOptions()
            options.add_experimental_option("detach", False)
            options.add_argument("--disable-blink-features=AutomationControlled")
            if self.headless:
                options.add_argument("--headless=new")
                options.add_argument("--window-size=1920,1080")
            driver = webdriver.Chrome(options=options)
        
        try:
            # Navigate to OWA first to set cookies on the right domain
            driver.get(f"{self.base_url}/owa/")
            time.sleep(2)
            
            # Add cookies (only ones that belong to this domain)
            for name, value in cookies.items():
                if name != "X-OWA-CANARY":  # Don't set the canary we don't have
                    try:
                        driver.add_cookie({"name": name, "value": value})
                    except Exception as e:
                        logger.debug(f"Could not add cookie {name}: {e}")
            
            # Navigate to calendar and wait for authentication
            print("🔗 Navigating to calendar...")
            driver.get(f"{self.base_url}/owa/?path=/calendar")
            
            # Wait for page to load
            WebDriverWait(driver, 10).until(
                lambda d: d.current_url != "data:," and d.current_url != "about:blank"
            )
            
            # Check if we're redirected to login
            if "login.microsoftonline" in driver.current_url:
                print("🔐 Please complete authentication in the browser...")
                # Wait for authentication to complete
                max_wait = 600
                start_time = time.time()
                while "login.microsoftonline" in driver.current_url:
                    if time.time() - start_time > max_wait:
                        driver.quit()
                        raise AuthenticationError("Authentication timeout")
                    time.sleep(2)
            
            print("⏳ Waiting for calendar to load...")
            max_wait = 300
            start_time = time.time()
            while "outlook.office.com" not in driver.current_url:
                if time.time() - start_time > max_wait:
                    raise AuthenticationError("Authentication timeout")
                time.sleep(2)
            
            # Store the driver for API calls
            self._driver = driver
            print("✅ Browser ready for API calls")
            
        except Exception as e:
            logger.error(f"Failed to launch browser for API: {e}")
            try:
                driver.quit()
            except:
                pass
            raise

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

    def fetch_calendar_events_via_browser(
        self,
        start_date: str,
        end_date: str,
    ) -> Optional[list[dict]]:
        """
        Fetch calendar events by making API calls from within the browser context.
        
        This is necessary for Office 365 where the OAuth tokens are not accessible
        as simple cookies.
        
        Args:
            start_date: Start date in ISO format (YYYY-MM-DDTHH:MM:SSZ)
            end_date: End date in ISO format (YYYY-MM-DDTHH:MM:SSZ)
            
        Returns:
            List of calendar events as dictionaries, or None if failed
        """
        # Reuse existing browser if available (from use_browser_api mode)
        if self._driver:
            print("📅 Using existing browser session for API calls...")
            driver = self._driver
            need_to_close = False
        else:
            headless_msg = " (headless)" if self.headless else ""
            print(f"🌐 Opening new browser{headless_msg} to fetch calendar events...")
            need_to_close = True

            # Create browser
            if self.browser == "edge":
                try:
                    from selenium.webdriver.edge.options import Options
                    options = Options()
                    options.add_experimental_option("detach", False)
                    options.add_argument("--disable-blink-features=AutomationControlled")
                    if self.headless:
                        options.add_argument("--headless=new")
                        options.add_argument("--window-size=1920,1080")
                    driver = webdriver.Edge(options=options)
                except ImportError:
                    raise AuthenticationError("Edge WebDriver not available")
            else:
                options = ChromeOptions()
                options.add_experimental_option("detach", False)
                options.add_argument("--disable-blink-features=AutomationControlled")
                if self.headless:
                    options.add_argument("--headless=new")
                    options.add_argument("--window-size=1920,1080")
                driver = webdriver.Chrome(options=options)
        
        try:
            # Navigate to OWA calendar if not already there
            current_url = driver.current_url
            if "calendar" not in current_url.lower():
                calendar_url = f"{self.base_url}/owa/?path=/calendar"
                print(f"🔗 Navigating to {calendar_url}...")
                driver.get(calendar_url)
            
            # Wait for page to load and check if we need to authenticate
            WebDriverWait(driver, 10).until(
                lambda d: d.current_url != "data:," and d.current_url != "about:blank"
            )
            
            # Check if we're redirected to login
            if "login.microsoftonline" in driver.current_url:
                print("🔐 Please complete authentication in the browser...")
                # Wait for authentication to complete
                max_wait = 600
                start_time = time.time()
                while "login.microsoftonline" in driver.current_url:
                    if time.time() - start_time > max_wait:
                        raise AuthenticationError("Authentication timeout")
                    time.sleep(2)
                
                # Wait for OWA to load
                time.sleep(5)
            
            # Wait for calendar page to be ready
            print("⏳ Waiting for calendar to load...")
        
            max_wait = 300
            start_time = time.time()
            while "outlook.office.com" not in driver.current_url:
                if time.time() - start_time > max_wait:
                    raise AuthenticationError("Authentication timeout")
                time.sleep(2)
            



            # Make the API call from within the browser using fetch
            print("📅 Fetching calendar events...")
            result = driver.execute_script(f"""
                return new Promise((resolve, reject) => {{
                    // Try Graph API first (works in OWA context)
                    fetch('https://outlook.office.com/api/v2.0/me/calendarview?startDateTime={start_date}&endDateTime={end_date}&$top=500&$select=Id,Subject,Start,End,Location,Organizer,Attendees,IsAllDay,IsCancelled,ShowAs,Body,Categories,Recurrence', {{
                        method: 'GET',
                        credentials: 'include',
                        headers: {{
                            'Content-Type': 'application/json',
                            'Accept': 'application/json'
                        }}
                    }})
                    .then(response => {{
                        console.log('REST API v2.0 status:', response.status);
                        if (response.ok) {{
                            return response.json().then(data => ({{status: response.status, data: data}}));
                        }} else {{
                            return response.text().then(text => ({{status: response.status, error: text}}));
                        }}
                    }})
                    .then(result => {{
                        if (result.data) {{
                            resolve(JSON.stringify({{success: true, events: result.data.value || [], source: 'rest_v2'}}));
                        }} else {{
                            resolve(JSON.stringify({{success: false, error: 'REST API status ' + result.status, detail: result.error}}));
                        }}
                    }})
                    .catch(error => {{
                        console.error('Calendar API error:', error);
                        resolve(JSON.stringify({{success: false, error: error.message}}));
                    }});
                    
                    // Timeout after 30 seconds
                    setTimeout(() => resolve(JSON.stringify({{success: false, error: 'timeout'}})), 30000);
                }});
            """)
            
            if result:
                parsed = json.loads(result)
                if parsed.get('success'):
                    events = parsed.get('events', [])
                    print(f"✅ Retrieved {len(events)} events via {parsed.get('source', 'unknown')}")
                    if need_to_close:
                        driver.quit()
                    return events
                else:
                    print(f"⚠️  REST API v2.0 failed: {parsed.get('error')}")
                    if parsed.get('detail'):
                        print(f"   Detail: {str(parsed.get('detail'))[:200]}")
            
            # Try alternative: Use OWA's internal API with more thorough canary search
            print("⏳ Trying OWA internal API...")
            result = driver.execute_script(f"""
                return new Promise((resolve, reject) => {{
                    (async function() {{
                    // Get the canary from the page - try MANY locations
                    let canary = '';
                    try {{
                        // Method 1: Common JS variables
                        if (window.g_CanaryValue) canary = window.g_CanaryValue;
                        else if (window.__owa_boot && window.__owa_boot.canary) canary = window.__owa_boot.canary;
                        
                        // Method 2: Cookies
                        if (!canary) {{
                            let cookies = document.cookie.split(';');
                            for (let c of cookies) {{
                                if (c.trim().startsWith('X-OWA-CANARY=')) {{
                                    canary = c.trim().split('=')[1];
                                    break;
                                }}
                            }}
                        }}
                        
                        // Method 3: sessionStorage
                        if (!canary) {{
                            for (let i = 0; i < sessionStorage.length; i++) {{
                                let key = sessionStorage.key(i);
                                let value = sessionStorage.getItem(key);
                                if (key.toLowerCase().includes('canary') || 
                                    (value && value.length > 20 && value.length < 200 && /^[a-zA-Z0-9_-]+$/.test(value))) {{
                                    canary = value;
                                    break;
                                }}
                            }}
                        }}
                        
                        // Method 4: Look in script tags
                        if (!canary) {{
                            let scripts = document.getElementsByTagName('script');
                            for (let i = 0; i < scripts.length; i++) {{
                                let content = scripts[i].textContent || '';
                                let match = content.match(/"canary"\\s*:\\s*"([^"]+)"/);
                                if (match) {{
                                    canary = match[1];
                                    break;
                                }}
                            }}
                        }}
                        
                        // Method 5: Try to get from network performance entries
                        if (!canary && window.performance) {{
                            let entries = performance.getEntriesByType('resource');
                            for (let entry of entries) {{
                                if (entry.name && entry.name.includes('X-OWA-CANARY=')) {{
                                    let match = entry.name.match(/X-OWA-CANARY=([^&]+)/);
                                    if (match) canary = match[1];
                                    break;
                                }}
                            }}
                        }}
                    }} catch(e) {{
                        console.error('Canary search error:', e);
                    }}
                    
                    console.log('Using canary:', canary ? canary.substring(0, 20) + '...' : 'none');
                    
                    // Try multiple OWA API endpoints (different paths for different Office 365 versions)
                    const endpoints = [
                        '/owa/0/service.svc?action=GetCalendarView',
                        '/mail/0/service.svc?action=GetCalendarView',
                        '/owa/service.svc?action=GetCalendarView',
                        '/mail/service.svc?action=GetCalendarView'
                    ];
                    
                    let lastError = null;
                    let lastStatus = null;
                    
                    async function tryEndpoint(endpoint) {{
                        try {{
                            console.log('Trying endpoint:', endpoint);
                            const response = await fetch(endpoint, {{
                                method: 'POST',
                                credentials: 'include',
                                headers: {{
                                    'Content-Type': 'application/json',
                                    'Action': 'GetCalendarView',
                                    'X-OWA-CANARY': canary,
                                    'X-Requested-With': 'XMLHttpRequest'
                                }},
                                body: JSON.stringify({{
                                    "__type": "GetCalendarViewJsonRequest:#Exchange",
                                    "Header": {{
                                        "__type": "JsonRequestHeaders:#Exchange",
                                        "RequestServerVersion": "Exchange2016"
                                    }},
                                    "Body": {{
                                        "__type": "GetCalendarViewRequest:#Exchange",
                                        "FolderId": {{
                                            "__type": "DistinguishedFolderId:#Exchange",
                                            "Id": "calendar"
                                        }},
                                        "StartDate": "{start_date}",
                                        "EndDate": "{end_date}"
                                    }}
                                }})
                            }});
                            
                            console.log('Endpoint', endpoint, 'status:', response.status);
                            lastStatus = response.status;
                            
                            if (response.ok) {{
                                const text = await response.text();
                                if (text && text.trim()) {{
                                    try {{
                                        return {{success: true, data: JSON.parse(text)}};
                                    }} catch(e) {{
                                        lastError = 'Invalid JSON';
                                    }}
                                }}
                            }} else {{
                                lastError = 'HTTP ' + response.status;
                            }}
                        }} catch(e) {{
                            console.error('Endpoint error:', e);
                            lastError = e.message;
                        }}
                        return {{success: false}};
                    }}
                    
                    // Try each endpoint
                    for (const endpoint of endpoints) {{
                        const result = await tryEndpoint(endpoint);
                        if (result.success) {{
                            let items = [];
                            try {{
                                const data = result.data;
                                if (data.Body && data.Body.Items) {{
                                    items = data.Body.Items;
                                }} else if (data.Body && data.Body.ResponseMessages) {{
                                    const msgs = data.Body.ResponseMessages.Items;
                                    if (msgs && msgs[0] && msgs[0].RootFolder) {{
                                        items = msgs[0].RootFolder.Items || [];
                                    }}
                                }}
                            }} catch(e) {{
                                console.error('Parse error:', e);
                            }}
                            resolve(JSON.stringify({{items: items, status: lastStatus, hasCanary: !!canary, endpoint: endpoint}}));
                            return;
                        }}
                    }}
                    
                    // All endpoints failed
                    resolve(JSON.stringify({{items: [], error: lastError, status: lastStatus, hasCanary: !!canary}}));
                    }})();  // End of async IIFE
                    
                    setTimeout(() => resolve(JSON.stringify({{items: [], error: 'timeout', hasCanary: false}})), 30000);
                }});
            """)
            
            if result:
                parsed = json.loads(result)
                events = parsed.get('items', [])
                error = parsed.get('error')
                has_canary = parsed.get('hasCanary', False)
                status = parsed.get('status', 'unknown')
                endpoint = parsed.get('endpoint', 'unknown')
                
                if error:
                    print(f"⚠️  OWA API error: {error} (status: {status}, hasCanary: {has_canary})")
                elif len(events) > 0:
                    print(f"✅ Retrieved {len(events)} events via OWA API ({endpoint})")
                    if need_to_close:
                        driver.quit()
                    return events
                else:
                    print(f"📅 OWA API returned 0 events (endpoint: {endpoint}, hasCanary: {has_canary})")
            
            # Method 3: Extract events from DOM using week view.
            # Week view shows all events per day without the truncation
            # that month view suffers from (month view silently hides overflow).
            from datetime import datetime as _dt, timedelta as _td

            print("⏳ Switching to week view...")
            try:
                driver.get(f"{self.base_url}/calendar/view/week")
                time.sleep(3)
                logger.info("Switched to week view for DOM extraction")
            except Exception as e:
                logger.warning(f"Failed to switch to week view: {e}")

            # JavaScript to extract events from the current calendar view
            extract_dom_js = """
                let events = [];
                try {
                    const selectors = [
                        '[role="button"][aria-label*="event"]',
                        '[role="button"][aria-label*=", "][aria-label*=" to "]',
                        '[data-automation-id="CalendarEventCard"]',
                        '.ms-CalendarEvent'
                    ];
                    let foundElements = new Set();
                    for (const selector of selectors) {
                        const elements = document.querySelectorAll(selector);
                        for (const el of elements) {
                            const label = el.getAttribute('aria-label');
                            if (label && label.length > 10 && label.includes(' to ')) {
                                if (foundElements.has(label)) continue;
                                foundElements.add(label);
                                const parts = label.split(', ');
                                if (parts.length >= 4) {
                                    const subject = parts[0];
                                    const timeRange = parts[1];
                                    let dayName = '';
                                    let monthDay = '';
                                    let year = '';
                                    let organizer = '';
                                    let isRecurring = label.includes('Recurring event');
                                    for (let i = 2; i < parts.length; i++) {
                                        const part = parts[i].trim();
                                        if (['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'].some(d => part.startsWith(d))) {
                                            dayName = part;
                                        } else if (['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'].some(m => part.startsWith(m))) {
                                            monthDay = part;
                                        } else if (/^\\d{4}$/.test(part)) {
                                            year = part;
                                        } else if (part.startsWith('By ')) {
                                            organizer = part.substring(3);
                                        }
                                    }
                                    let startTime = '';
                                    let endTime = '';
                                    function to24h(time, period) {
                                        if (!time) return '';
                                        let [h, m] = time.split(':').map(Number);
                                        if (period && period.toUpperCase() === 'PM' && h !== 12) h += 12;
                                        if (period && period.toUpperCase() === 'AM' && h === 12) h = 0;
                                        return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
                                    }
                                    const timeMatch12 = timeRange.match(/(\\d{1,2}:\\d{2})\\s*(AM|PM)?\\s+to\\s+(\\d{1,2}:\\d{2})\\s*(AM|PM)?/i);
                                    if (timeMatch12) {
                                        let startPeriod = timeMatch12[2] || timeMatch12[4] || 'PM';
                                        let endPeriod = timeMatch12[4] || startPeriod;
                                        startTime = to24h(timeMatch12[1], startPeriod);
                                        endTime = to24h(timeMatch12[3], endPeriod);
                                    } else {
                                        const timeMatch24 = timeRange.match(/(\\d{1,2}:\\d{2})\\s+to\\s+(\\d{1,2}:\\d{2})/);
                                        if (timeMatch24) {
                                            startTime = timeMatch24[1].padStart(5, '0');
                                            endTime = timeMatch24[2].padStart(5, '0');
                                        }
                                    }
                                    let startDate = null;
                                    let endDate = null;
                                    if (monthDay && year) {
                                        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                                                           'July', 'August', 'September', 'October', 'November', 'December'];
                                        const monthMatch = monthDay.match(/([A-Za-z]+)\\s+(\\d+)/);
                                        if (monthMatch) {
                                            const monthIdx = monthNames.findIndex(m => m.toLowerCase() === monthMatch[1].toLowerCase());
                                            const day = parseInt(monthMatch[2]);
                                            if (monthIdx >= 0 && day > 0) {
                                                const datePrefix = year + '-' + String(monthIdx + 1).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                                                startDate = datePrefix + 'T' + (startTime || '00:00') + ':00';
                                                endDate = datePrefix + 'T' + (endTime || '23:59') + ':00';
                                            }
                                        }
                                    }
                                    events.push({
                                        Subject: subject,
                                        Start: startDate ? {DateTime: startDate, TimeZone: 'UTC'} : null,
                                        End: endDate ? {DateTime: endDate, TimeZone: 'UTC'} : null,
                                        Organizer: organizer ? {EmailAddress: {Name: organizer}} : null,
                                        IsRecurring: isRecurring,
                                        _rawLabel: label,
                                        _source: 'dom'
                                    });
                                }
                            }
                        }
                    }
                } catch(e) {
                    console.error('DOM extraction error:', e);
                }
                return JSON.stringify(events);
            """

            # Calculate how many weeks to navigate forward
            try:
                start_dt = _dt.strptime(start_date[:10], "%Y-%m-%d")
                end_dt = _dt.strptime(end_date[:10], "%Y-%m-%d")
            except (ValueError, TypeError):
                start_dt = _dt.now()
                end_dt = start_dt + _td(days=15)

            now = _dt.now()
            current_week_end = now + _td(days=(6 - now.weekday()))  # Sunday of current week
            weeks_forward = 0
            while current_week_end < end_dt:
                weeks_forward += 1
                current_week_end += _td(days=7)

            total_weeks = 1 + weeks_forward
            all_dom_events = []
            seen_labels = set()

            print(f"⏳ Extracting events from {total_weeks} week view(s)...")

            # Extract current week
            dom_events = driver.execute_script(extract_dom_js)
            if dom_events:
                week_parsed = json.loads(dom_events)
                for event in week_parsed:
                    raw = event.get("_rawLabel", "")
                    if raw not in seen_labels:
                        seen_labels.add(raw)
                        all_dom_events.append(event)
                logger.info(f"  Week 1/{total_weeks} (current): {len(all_dom_events)} events")

            # Click forward through remaining weeks
            for w in range(weeks_forward):
                try:
                    # Find the "Go to next week" button in the main calendar view
                    next_btn = None
                    try:
                        next_btn = driver.find_element("css selector", 'button[aria-label*="next week" i]')
                    except Exception:
                        pass

                    if not next_btn:
                        logger.warning(f"  Week {w+2}/{total_weeks}: could not find next-week button")
                        break

                    next_btn.click()
                    time.sleep(4)

                    dom_events = driver.execute_script(extract_dom_js)
                    if dom_events:
                        week_parsed = json.loads(dom_events)
                        added = 0
                        for event in week_parsed:
                            raw = event.get("_rawLabel", "")
                            if raw not in seen_labels:
                                seen_labels.add(raw)
                                all_dom_events.append(event)
                                added += 1
                        logger.info(f"  Week {w+2}/{total_weeks}: {added} new events ({len(week_parsed)} total, {len(week_parsed) - added} duplicates)")
                    else:
                        logger.info(f"  Week {w+2}/{total_weeks}: no events found")
                except Exception as e:
                    logger.warning(f"  Week {w+2}/{total_weeks}: navigation failed: {e}")

            if len(all_dom_events) > 0:
                logger.info(f"Total DOM events extracted across all weeks: {len(all_dom_events)}")
                if need_to_close:
                    driver.quit()
                for event in all_dom_events:
                    logger.debug(f"  {event['Subject']} @ {event['Start']['DateTime'] if event.get('Start') else '?'}")
                return all_dom_events
            
            # Only close browser if we opened it ourselves (not in use_browser_api mode)
            if need_to_close:
                driver.quit()
            
            print("❌ Could not retrieve calendar events via any method")
            return []
            
        except Exception as e:
            logger.error(f"Browser calendar fetch failed: {e}")
            if need_to_close:
                try:
                    driver.quit()
                except:
                    pass
            return None
