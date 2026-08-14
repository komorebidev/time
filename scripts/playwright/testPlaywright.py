from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")

    context = browser.contexts[0]

    print(f"Found {len(context.pages)} open tabs:")

    for page in context.pages:
        print(f"  {page.title()}")
        print(f"  {page.url}")
