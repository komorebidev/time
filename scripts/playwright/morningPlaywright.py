from playwright.sync_api import sync_playwright

BASE_URL = "https://support.eiresystems.com/ticket"

WORKLOG_TEXT = "Eire: Email catchup, internal communication, time recording, work logs"
STATUS = "Completed (On Hold)"
CHARGE_TYPE = "Internal work"

START_TIME = "09:00"
END_TIME = "10:00"


def main():
    ticket = input("Ticket number: ").strip()

    if not ticket.isdigit():
        print("Invalid ticket number.")
        return

    url = f"{BASE_URL}?id={ticket}&showalltickettypes=1"

    with sync_playwright() as p:
        # Connect to the already-open Playwright Edge session.
        browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")

        pages = browser.contexts[0].pages

        if not pages:
            print("No Edge pages found.")
            return

        page = pages[0]

        print(f"Opening ticket {ticket}...")
        page.goto(url)
        page.wait_for_load_state("domcontentloaded")

        print("Opening Worklog...")
        page.get_by_role("button", name="Worklog").click()

        # Worklog editor
        print("Entering worklog...")
        editor = page.locator('[contenteditable="true"]').first
        editor.click()
        editor.fill(WORKLOG_TEXT)

        # Status
        print(f"Setting status: {STATUS}")
        status = page.get_by_role("combobox", name="Status *")
        status.click()
        page.get_by_text(STATUS, exact=True).click()

        # Job Start time
        print(f"Setting start time: {START_TIME}")
        time_fields = page.locator('input').filter(has=page.locator(""))