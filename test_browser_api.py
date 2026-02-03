"""
Test script to manually authenticate and try various API calls.
This helps identify which endpoints work in the browser context.
"""

import time
import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait

def main():
    print("🌐 Opening browser...")
    
    options = ChromeOptions()
    options.add_experimental_option("detach", False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=options)
    
    try:
        # Navigate to OWA
        print("🔗 Navigating to Outlook Web App...")
        driver.get("https://outlook.office.com/calendar")
        
        # Wait for user to authenticate
        print("\n" + "="*60)
        print("🔐 Please complete authentication in the browser.")
        print("   When you see your calendar, press ENTER here...")
        print("="*60 + "\n")
        input()
        
        # Give it a moment to fully load
        print("⏳ Waiting for page to stabilize...")
        time.sleep(3)
        
        # Get current URL
        print(f"📍 Current URL: {driver.current_url}")
        
        # Try to find the canary token
        print("\n" + "="*60)
        print("🔍 Looking for X-OWA-CANARY token...")
        print("="*60)
        
        canary_result = driver.execute_script("""
            let results = {};
            
            // Check cookies
            results.cookies = document.cookie;
            
            // Check common JS variables
            results.g_CanaryValue = window.g_CanaryValue || null;
            results.owa_boot_canary = (window.__owa_boot && window.__owa_boot.canary) || null;
            
            // Check sessionStorage
            results.sessionStorage = {};
            for (let i = 0; i < sessionStorage.length; i++) {
                let key = sessionStorage.key(i);
                results.sessionStorage[key] = sessionStorage.getItem(key);
            }
            
            // Check localStorage
            results.localStorage = {};
            for (let i = 0; i < localStorage.length; i++) {
                let key = localStorage.key(i);
                if (key.toLowerCase().includes('canary') || key.toLowerCase().includes('token')) {
                    results.localStorage[key] = localStorage.getItem(key);
                }
            }
            
            // Look in script tags
            let scripts = document.getElementsByTagName('script');
            for (let i = 0; i < scripts.length; i++) {
                let content = scripts[i].textContent || '';
                let match = content.match(/"canary"\\s*:\\s*"([^"]+)"/);
                if (match) {
                    results.scriptTagCanary = match[1];
                    break;
                }
            }
            
            return JSON.stringify(results, null, 2);
        """)
        
        parsed = json.loads(canary_result)
        print(f"\nCookies: {parsed.get('cookies', '')[:200]}...")
        print(f"g_CanaryValue: {parsed.get('g_CanaryValue')}")
        print(f"__owa_boot.canary: {parsed.get('owa_boot_canary')}")
        print(f"scriptTagCanary: {parsed.get('scriptTagCanary')}")
        print(f"sessionStorage keys: {list(parsed.get('sessionStorage', {}).keys())}")
        print(f"localStorage keys: {list(parsed.get('localStorage', {}).keys())}")
        
        # Extract canary for use in API calls
        canary = (
            parsed.get('g_CanaryValue') or 
            parsed.get('owa_boot_canary') or 
            parsed.get('scriptTagCanary') or
            ''
        )
        
        # Check cookies for canary
        if not canary:
            cookies = parsed.get('cookies', '')
            if 'X-OWA-CANARY=' in cookies:
                for part in cookies.split(';'):
                    if 'X-OWA-CANARY=' in part:
                        canary = part.split('=', 1)[1].strip()
                        break
        
        print(f"\n✅ Canary found: {canary[:50] if canary else 'NOT FOUND'}...")
        
        # Look for MSAL access tokens in localStorage
        print("\n" + "="*60)
        print("🔍 Looking for MSAL access tokens in localStorage...")
        print("="*60)
        
        msal_tokens = driver.execute_script("""
            let tokens = {};
            for (let i = 0; i < localStorage.length; i++) {
                let key = localStorage.key(i);
                if (key.includes('accesstoken') && key.includes('outlook.office.com')) {
                    try {
                        let value = localStorage.getItem(key);
                        let parsed = JSON.parse(value);
                        tokens[key.substring(0, 100)] = {
                            keys: Object.keys(parsed),
                            hasSecret: !!parsed.secret,
                            secretPreview: parsed.secret ? parsed.secret.substring(0, 50) + '...' : null,
                            hasCredentialType: !!parsed.credentialType,
                            expiresOn: parsed.expiresOn,
                            target: parsed.target ? parsed.target.substring(0, 100) : null,
                            // Show full structure
                            fullValue: JSON.stringify(parsed).substring(0, 300)
                        };
                    } catch(e) {
                        tokens[key.substring(0, 50)] = 'parse error: ' + e.message;
                    }
                }
            }
            return JSON.stringify(tokens, null, 2);
        """)
        print(msal_tokens)
        
        # Try to get ANY access token with secret
        print("\n" + "="*60)
        print("🔍 Looking for ANY token with secret field...")
        print("="*60)
        
        all_tokens_info = driver.execute_script("""
            let results = [];
            for (let i = 0; i < localStorage.length; i++) {
                let key = localStorage.key(i);
                try {
                    let value = localStorage.getItem(key);
                    let parsed = JSON.parse(value);
                    if (parsed && parsed.secret) {
                        results.push({
                            key: key.substring(0, 80),
                            secretLength: parsed.secret.length,
                            secretPreview: parsed.secret.substring(0, 30) + '...',
                            target: parsed.target || 'no target'
                        });
                    }
                } catch(e) {}
            }
            return JSON.stringify(results, null, 2);
        """)
        print(all_tokens_info)
        
        # Extract the actual access token for outlook.office.com
        access_token = driver.execute_script("""
            for (let i = 0; i < localStorage.length; i++) {
                let key = localStorage.key(i);
                if (key.includes('accesstoken') && key.includes('outlook.office.com')) {
                    try {
                        let value = localStorage.getItem(key);
                        let parsed = JSON.parse(value);
                        if (parsed.secret) {
                            return parsed.secret;
                        }
                    } catch(e) {}
                }
            }
            return null;
        """)
        
        if access_token:
            print(f"\n✅ Found access token: {access_token[:50]}...")
        else:
            print("\n❌ No access token found for calendars scope")
        
        # Calculate date range
        from datetime import datetime, timedelta
        start = datetime.now()
        end = start + timedelta(days=7)
        start_str = start.strftime("%Y-%m-%dT00:00:00Z")
        end_str = end.strftime("%Y-%m-%dT23:59:59Z")
        
        print(f"\n📅 Date range: {start_str} to {end_str}")
        
        # Test various API endpoints
        print("\n" + "="*60)
        print("🧪 Testing API endpoints...")
        print("="*60)
        
        endpoints_to_test = []
        
        # If we have an access token, test with Bearer auth first
        if access_token:
            endpoints_to_test.extend([
                # Outlook REST API with Bearer token
                {
                    "name": "Outlook REST API v2.0 with Bearer token",
                    "method": "GET",
                    "url": f"https://outlook.office.com/api/v2.0/me/calendarview?startDateTime={start_str}&endDateTime={end_str}&$top=50",
                    "headers": {
                        "Accept": "application/json",
                        "Authorization": f"Bearer {access_token}"
                    }
                },
                # Outlook REST API - calendars endpoint
                {
                    "name": "Outlook REST API calendars with Bearer",
                    "method": "GET",
                    "url": "https://outlook.office.com/api/v2.0/me/calendars",
                    "headers": {
                        "Accept": "application/json",
                        "Authorization": f"Bearer {access_token}"
                    }
                },
                # Outlook REST API - events endpoint
                {
                    "name": "Outlook REST API events with Bearer",
                    "method": "GET",
                    "url": f"https://outlook.office.com/api/v2.0/me/events?$top=10",
                    "headers": {
                        "Accept": "application/json",
                        "Authorization": f"Bearer {access_token}"
                    }
                },
            ])
        
        # Also test without Bearer (cookie-based)
        endpoints_to_test.extend([
            # Graph API via outlook.office.com
            {
                "name": "Graph API (via outlook.office.com) - no Bearer",
                "method": "GET",
                "url": f"https://outlook.office.com/api/v2.0/me/calendarview?startDateTime={start_str}&endDateTime={end_str}",
                "headers": {"Accept": "application/json"}
            },
            # Graph API direct
            {
                "name": "Microsoft Graph API direct",
                "method": "GET", 
                "url": f"https://graph.microsoft.com/v1.0/me/calendarview?startDateTime={start_str}&endDateTime={end_str}",
                "headers": {"Accept": "application/json"}
            },
            # OWA Calendar endpoint (modern)
            {
                "name": "OWA Calendar (modern path /mail/)",
                "method": "POST",
                "url": "/mail/0/service.svc?action=GetCalendarView",
                "headers": {
                    "Content-Type": "application/json",
                    "Action": "GetCalendarView",
                    "X-OWA-CANARY": canary
                },
                "body": {
                    "__type": "GetCalendarViewJsonRequest:#Exchange",
                    "Header": {
                        "__type": "JsonRequestHeaders:#Exchange", 
                        "RequestServerVersion": "Exchange2016"
                    },
                    "Body": {
                        "__type": "GetCalendarViewRequest:#Exchange",
                        "FolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": "calendar"},
                        "StartDate": start_str,
                        "EndDate": end_str
                    }
                }
            },
            # OWA service.svc classic
            {
                "name": "OWA service.svc (classic /owa/)",
                "method": "POST",
                "url": "/owa/service.svc?action=GetCalendarView",
                "headers": {
                    "Content-Type": "application/json",
                    "Action": "GetCalendarView", 
                    "X-OWA-CANARY": canary
                },
                "body": {
                    "__type": "GetCalendarViewJsonRequest:#Exchange",
                    "Header": {
                        "__type": "JsonRequestHeaders:#Exchange",
                        "RequestServerVersion": "Exchange2016"
                    },
                    "Body": {
                        "__type": "GetCalendarViewRequest:#Exchange",
                        "FolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": "calendar"},
                        "StartDate": start_str,
                        "EndDate": end_str
                    }
                }
            },
            # OWA FindItem
            {
                "name": "OWA FindItem",
                "method": "POST",
                "url": "/owa/service.svc?action=FindItem",
                "headers": {
                    "Content-Type": "application/json",
                    "Action": "FindItem",
                    "X-OWA-CANARY": canary
                },
                "body": {
                    "__type": "FindItemJsonRequest:#Exchange",
                    "Header": {
                        "__type": "JsonRequestHeaders:#Exchange",
                        "RequestServerVersion": "Exchange2016"
                    },
                    "Body": {
                        "__type": "FindItemRequest:#Exchange",
                        "ItemShape": {"__type": "ItemResponseShape:#Exchange", "BaseShape": "Default"},
                        "ParentFolderIds": [{"__type": "DistinguishedFolderId:#Exchange", "Id": "calendar"}],
                        "Traversal": "Shallow"
                    }
                }
            },
            # Try /calendar/ path
            {
                "name": "OWA /calendar/ service",
                "method": "POST",
                "url": "/calendar/0/service.svc?action=GetCalendarView",
                "headers": {
                    "Content-Type": "application/json",
                    "Action": "GetCalendarView",
                    "X-OWA-CANARY": canary
                },
                "body": {
                    "__type": "GetCalendarViewJsonRequest:#Exchange",
                    "Header": {"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": "Exchange2016"},
                    "Body": {
                        "__type": "GetCalendarViewRequest:#Exchange",
                        "FolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": "calendar"},
                        "StartDate": start_str,
                        "EndDate": end_str
                    }
                }
            }
        ])
        
        for endpoint in endpoints_to_test:
            print(f"\n📡 Testing: {endpoint['name']}")
            print(f"   URL: {endpoint['url'][:80]}...")
            
            # Build fetch options
            fetch_options = {
                "method": endpoint["method"],
                "credentials": "include",
                "headers": endpoint["headers"]
            }
            if endpoint.get("body"):
                fetch_options["body"] = json.dumps(endpoint["body"])
            
            js_code = f"""
                return fetch("{endpoint['url']}", {json.dumps(fetch_options)})
                    .then(response => {{
                        return response.text().then(text => ({{
                            status: response.status,
                            statusText: response.statusText,
                            ok: response.ok,
                            body: text.substring(0, 500)
                        }}));
                    }})
                    .catch(error => ({{
                        error: error.message
                    }}));
            """
            
            try:
                result = driver.execute_script(js_code)
                
                if result.get("error"):
                    print(f"   ❌ Error: {result['error']}")
                elif result.get("ok"):
                    print(f"   ✅ Status: {result['status']} {result['statusText']}")
                    body = result.get('body', '')
                    if body:
                        try:
                            parsed_body = json.loads(body)
                            if isinstance(parsed_body, dict):
                                if 'value' in parsed_body:
                                    print(f"   📅 Found {len(parsed_body['value'])} events!")
                                elif 'Body' in parsed_body:
                                    print(f"   📦 Got Body response")
                                else:
                                    print(f"   📦 Keys: {list(parsed_body.keys())[:5]}")
                        except:
                            print(f"   📄 Response: {body[:100]}...")
                else:
                    print(f"   ⚠️  Status: {result['status']} {result['statusText']}")
                    print(f"   📄 Response: {result.get('body', '')[:200]}")
                    
            except Exception as e:
                print(f"   ❌ Exception: {e}")
        
        print("\n" + "="*60)
        print("🔍 Checking what OWA React app exposes...")
        print("="*60)
        
        # Try to find calendar data in the React app state
        react_check = driver.execute_script("""
            let results = {};
            
            // Check for exposed APIs
            results.hasGraphQL = !!window.__APOLLO_STATE__;
            results.hasReactRoot = !!document.querySelector('#app')?._reactRootContainer;
            
            // Look for calendar-related global objects
            let calendarObjects = [];
            for (let key in window) {
                if (key.toLowerCase().includes('calendar') || 
                    key.toLowerCase().includes('event') ||
                    key.toLowerCase().includes('owa')) {
                    calendarObjects.push(key);
                }
            }
            results.calendarObjects = calendarObjects.slice(0, 20);
            
            // Check for modern OWA boot config
            if (window.__owa_boot) {
                results.owaBootKeys = Object.keys(window.__owa_boot);
            }
            
            return JSON.stringify(results, null, 2);
        """)
        
        print(react_check)
        
        # Try to use OWA's internal mechanisms
        print("\n" + "="*60)
        print("🔍 Exploring OWA's internal API mechanism...")
        print("="*60)
        
        owa_internals = driver.execute_script("""
            let results = {};
            
            // Check if Owa object exists and what's in it
            if (window.Owa) {
                results.OwaKeys = Object.keys(window.Owa).slice(0, 30);
            }
            
            // Check webpackChunkOwa for exposed modules
            if (window.webpackChunkOwa && window.webpackChunkOwa.length > 0) {
                results.webpackChunks = window.webpackChunkOwa.length;
            }
            
            // Try to find the request infrastructure
            let potentialApis = [];
            for (let key in window) {
                let val = window[key];
                if (val && typeof val === 'object') {
                    if (val.fetch || val.request || val.post || val.get) {
                        potentialApis.push(key);
                    }
                }
            }
            results.potentialApis = potentialApis.slice(0, 10);
            
            // Look for headers that OWA uses
            results.owaHeaders = {};
            let meta = document.querySelector('meta[name="ms.owaconfig"]');
            if (meta) {
                results.owaConfig = meta.content.substring(0, 200);
            }
            
            return JSON.stringify(results, null, 2);
        """)
        print(owa_internals)
        
        # Try using the Owa global object if available
        print("\n" + "="*60)
        print("🧪 Trying OWA's built-in API functions...")
        print("="*60)
        
        owa_api_test = driver.execute_script("""
            return new Promise((resolve) => {
                let results = {};
                
                // Try to access OWA's service worker or API layer
                if (window.Owa && window.Owa.Service) {
                    results.hasOwaService = true;
                    results.OwaServiceKeys = Object.keys(window.Owa.Service);
                }
                
                // Check for exposed action creators or stores (Redux-like)
                if (window.__REDUX_DEVTOOLS_EXTENSION__) {
                    results.hasRedux = true;
                }
                
                // Try to find calendar store/state
                if (window.Owa && window.Owa.stores) {
                    results.storeKeys = Object.keys(window.Owa.stores);
                }
                
                // Look for exposed calendar data in Owa namespace
                function searchObject(obj, depth, path) {
                    if (depth > 3 || !obj || typeof obj !== 'object') return null;
                    
                    for (let key of Object.keys(obj).slice(0, 20)) {
                        if (key.toLowerCase().includes('calendar') || key.toLowerCase().includes('event')) {
                            return {path: path + '.' + key, type: typeof obj[key]};
                        }
                        if (typeof obj[key] === 'object' && obj[key] !== null) {
                            let found = searchObject(obj[key], depth + 1, path + '.' + key);
                            if (found) return found;
                        }
                    }
                    return null;
                }
                
                if (window.Owa) {
                    results.calendarInOwa = searchObject(window.Owa, 0, 'Owa');
                }
                
                resolve(JSON.stringify(results, null, 2));
            });
        """)
        print(owa_api_test)
        
        # Try intercepting network requests by hooking fetch
        print("\n" + "="*60)
        print("🔍 Looking for how OWA makes authenticated requests...")
        print("="*60)
        
        network_info = driver.execute_script("""
            // Check performance entries for API calls
            let entries = performance.getEntriesByType('resource');
            let apiCalls = entries.filter(e => 
                e.name.includes('service.svc') || 
                e.name.includes('/api/') ||
                e.name.includes('/owa/')
            ).slice(-10);
            
            return JSON.stringify(apiCalls.map(e => ({
                url: e.name.substring(0, 100),
                duration: e.duration
            })), null, 2);
        """)
        print("Recent API calls from performance entries:")
        print(network_info)

        print("\n" + "="*60)
        print("✅ Test complete! Press ENTER to close browser...")
        print("="*60)
        input()
        
    finally:
        driver.quit()
        print("👋 Browser closed.")

if __name__ == "__main__":
    main()
