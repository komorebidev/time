from playwright.sync_api import sync_playwright


with sync_playwright() as p:
    print("Connecting to Edge...")

    browser = p.chromium.connect_over_cdp(
        "http://127.0.0.1:9222"
    )

    print("Connected!")

    for i, context in enumerate(browser.contexts):
        print(f"\nContext {i}:")

        for j, page in enumerate(context.pages):
            print(f"  Tab {j}")
            print(f"    URL:   {page.url}")
            print(f"    Title: {page.title()}")