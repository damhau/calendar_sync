"""
Test script to intercept OWA's actual network requests and see how it authenticates.
"""

import time
import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By

def main():
    print("🌐 Opening browser...")
    
    options = ChromeOptions()
    options.add_experimental_option("detach", False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    # Enable performance logging
    options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
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
        
        print("⏳ Waiting for page to stabilize...")
        time.sleep(3)
        
        # Install network interceptor
        print("\n" + "="*60)
        print("🔍 Setting up network interception...")
        print("="*60)
        
        driver.execute_script("""
            // Store original fetch
            window._originalFetch = window.fetch;
            window._capturedRequests = [];
            
            // Override fetch to capture requests
            window.fetch = function(...args) {
                const [url, options] = args;
                const captured = {
                    url: typeof url === 'string' ? url : url.url,
                    method: options?.method || 'GET',
                    headers: {},
                    timestamp: new Date().toISOString()
                };
                
                // Capture headers
                if (options?.headers) {
                    if (options.headers instanceof Headers) {
                        options.headers.forEach((value, key) => {
                            captured.headers[key] = value.substring(0, 100);
                        });
                    } else {
                        for (let key in options.headers) {
                            captured.headers[key] = String(options.headers[key]).substring(0, 100);
                        }
                    }
                }
                
                window._capturedRequests.push(captured);
                console.log('Captured fetch:', captured.url);
                
                return window._originalFetch.apply(this, args);
            };
            
            console.log('Network interception installed');
        """)
        
        print("✅ Network interception installed")
        
        # Trigger calendar navigation to capture requests
        print("\n" + "="*60)
        print("🔄 Triggering calendar refresh...")
        print("="*60)
        
        # Navigate away and back to trigger data fetch
        driver.get("https://outlook.office.com/mail")
        time.sleep(2)
        driver.get("https://outlook.office.com/calendar")
        time.sleep(5)
        
        # Get captured requests
        captured = driver.execute_script("return window._capturedRequests || [];")
        
        print(f"\n📊 Captured {len(captured)} requests:")
        for req in captured:
            url = req.get('url', '')[:80]
            method = req.get('method', 'GET')
            headers = req.get('headers', {})
            
            # Only show API-related requests
            if 'service.svc' in url or '/api/' in url or 'graph' in url:
                print(f"\n   📡 {method} {url}")
                if headers:
                    print(f"   Headers: {json.dumps(headers, indent=6)}")
        
        # Also look at what's in the rendered calendar
        print("\n" + "="*60)
        print("📅 Looking for calendar events in the DOM...")
        print("="*60)
        
        events_from_dom = driver.execute_script("""
            let events = [];
            
            // Look for calendar event elements
            // OWA uses various class names for events
            const selectors = [
                '[data-automation-id="CalendarEventCard"]',
                '[role="button"][aria-label*="event"]',
                '.ms-CalendarEvent',
                '[class*="calendarEvent"]',
                '[class*="CalendarEvent"]',
                '[data-is-focusable="true"][aria-label]'
            ];
            
            for (const selector of selectors) {
                const elements = document.querySelectorAll(selector);
                for (const el of elements) {
                    const label = el.getAttribute('aria-label') || el.innerText;
                    if (label && label.length > 3 && label.length < 500) {
                        events.push({
                            selector: selector,
                            text: label.substring(0, 200)
                        });
                    }
                }
            }
            
            return events;
        """)
        
        print(f"Found {len(events_from_dom)} potential events in DOM:")
        for evt in events_from_dom[:10]:
            print(f"   [{evt['selector']}] {evt['text'][:100]}")
        
        # Try to find if there's a global data store
        print("\n" + "="*60)
        print("🔍 Looking for calendar data in app state...")
        print("="*60)
        
        app_state = driver.execute_script("""
            let results = {};
            
            // Check for webpackChunkOwa modules
            if (window.webpackChunkOwa) {
                results.chunkCount = window.webpackChunkOwa.length;
                
                // Try to find exported modules
                for (let i = 0; i < Math.min(10, window.webpackChunkOwa.length); i++) {
                    let chunk = window.webpackChunkOwa[i];
                    if (chunk && chunk[1]) {
                        let moduleKeys = Object.keys(chunk[1]).slice(0, 5);
                        results['chunk_' + i] = moduleKeys;
                    }
                }
            }
            
            // Check __INITIAL_STATE__ or similar
            if (window.__INITIAL_STATE__) {
                results.hasInitialState = true;
                results.initialStateKeys = Object.keys(window.__INITIAL_STATE__);
            }
            
            // Check for any global stores
            for (let key of ['store', 'Store', 'appStore', 'calendarStore', 'eventStore']) {
                if (window[key]) {
                    results['has_' + key] = true;
                }
            }
            
            return results;
        """)
        
        print(json.dumps(app_state, indent=2))
        
        # Try a different approach - use CDP to get network logs
        print("\n" + "="*60)
        print("🔍 Getting network logs via Chrome DevTools Protocol...")
        print("="*60)
        
        try:
            logs = driver.get_log('performance')
            api_requests = []
            
            for entry in logs:
                try:
                    message = json.loads(entry['message'])
                    if message.get('message', {}).get('method') == 'Network.requestWillBeSent':
                        params = message['message'].get('params', {})
                        request = params.get('request', {})
                        url = request.get('url', '')
                        
                        if 'service.svc' in url or '/api/' in url:
                            api_requests.append({
                                'url': url[:100],
                                'method': request.get('method'),
                                'headers': {k: v[:50] for k, v in list(request.get('headers', {}).items())[:10]}
                            })
                except:
                    pass
            
            print(f"Found {len(api_requests)} API requests in network logs:")
            for req in api_requests[:10]:
                print(f"\n   📡 {req['method']} {req['url']}")
                if 'Authorization' in req.get('headers', {}):
                    print(f"   🔑 Has Authorization header!")
                    print(f"   Authorization: {req['headers']['Authorization'][:80]}...")
                if 'X-OWA-CANARY' in req.get('headers', {}):
                    print(f"   🎫 X-OWA-CANARY: {req['headers']['X-OWA-CANARY'][:50]}...")
                    
        except Exception as e:
            print(f"Could not get performance logs: {e}")
        
        print("\n" + "="*60)
        print("✅ Test complete! Press ENTER to close browser...")
        print("="*60)
        input()
        
    finally:
        driver.quit()
        print("👋 Browser closed.")

if __name__ == "__main__":
    main()
