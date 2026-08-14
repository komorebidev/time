import os
import shutil
import subprocess
import sys
import time


# ============================================================
# Configuration
# ============================================================

BASE_URL = "https://support.eiresystems.com/ticket"

SESSION = "halo"

WORKLOG_TEXT = (
    "Eire: Email catchup, internal communication, "
    "time recording, work logs"
)

STATUS = "Completed (On Hold)"
CHARGE_TYPE = "Internal work"

# Hard-coded for this version
START_TIME = "09:00"
END_TIME = "10:00"


# ============================================================
# Find playwright-cli
# ============================================================

def find_playwright_cli():
    cli = shutil.which("playwright-cli")

    if cli:
        return cli

    # Fallback for Windows npm global installs
    appdata = os.environ.get("APPDATA")

    if appdata:
        candidates = [
            os.path.join(appdata, "npm", "playwright-cli.cmd"),
            os.path.join(appdata, "npm", "playwright-cli.CMD"),
            os.path.join(appdata, "npm", "playwright-cli"),
        ]

        for candidate in candidates:
            if os.path.isfile(candidate):
                return candidate

    raise FileNotFoundError(
        "Could not find playwright-cli. "
        "Make sure it is installed and available on PATH."
    )


PLAYWRIGHT_CLI = find_playwright_cli()


# ============================================================
# Run playwright-cli
# ============================================================

def pw(*args, check=True):
    """
    Run a playwright-cli command.

    UTF-8 decoding is forced because Windows may otherwise try
    to decode Playwright's output using the system code page.
    """

    command = [
        PLAYWRIGHT_CLI,
        f"--s={SESSION}",
        *args,
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    if result.stdout:
        print(result.stdout)

    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if check and result.returncode != 0:
        raise RuntimeError(
            "Playwright CLI command failed:\n"
            + " ".join(command)
            + f"\n\nExit code: {result.returncode}"
        )

    return result


# ============================================================
# Attach to Edge
# ============================================================

def attach():
    print("Attaching to Microsoft Edge...")

    # The CLI creates the "halo" session.
    result = subprocess.run(
        [
            PLAYWRIGHT_CLI,
            "attach",
            "--extension=msedge",
            f"--session={SESSION}",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    if result.stdout:
        print(result.stdout)

    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        raise RuntimeError(
            "Could not attach to Microsoft Edge.\n"
            + (result.stderr or result.stdout)
        )

    # Give the CLI session a moment to finish attaching.
    time.sleep(1)


# ============================================================
# Main automation
# ============================================================

def main():

    print(f"Using playwright-cli: {PLAYWRIGHT_CLI}")
    print()

    # --------------------------------------------------------
    # Attach to the already-open Edge session
    # --------------------------------------------------------

    attach()

    # --------------------------------------------------------
    # Ask for ticket
    # --------------------------------------------------------

    ticket = input("Ticket number: ").strip()

    if not ticket.isdigit():
        print("Invalid ticket number.")
        return

    url = f"{BASE_URL}?id={ticket}&showalltickettypes=1"

    # --------------------------------------------------------
    # Navigate
    # --------------------------------------------------------

    print()
    print(f"Opening ticket {ticket}...")

    pw("goto", url)

    # --------------------------------------------------------
    # Snapshot before Worklog
    # --------------------------------------------------------

    print()
    print("Taking snapshot...")

    pw("snapshot")

    # --------------------------------------------------------
    # Open Worklog
    # --------------------------------------------------------

    print()
    print("Opening Worklog...")

    pw(
        "click",
        "getByRole('button', { name: 'Worklog' })"
    )

    # Give HaloPSA a moment to render the Worklog form.
    time.sleep(1)

    # --------------------------------------------------------
    # Snapshot Worklog form
    # --------------------------------------------------------

    print()
    print("Reading Worklog form...")

    pw("snapshot")

    # --------------------------------------------------------
    # Enter Worklog text
    # --------------------------------------------------------

    print()
    print("Entering worklog...")

    pw(
        "fill",
        "locator('[contenteditable=\"true\"]').first()",
        WORKLOG_TEXT,
    )

    # --------------------------------------------------------
    # Set Status
    # --------------------------------------------------------

    print()
    print(f"Setting status: {STATUS}")

    pw(
        "select",
        "getByRole('combobox', { name: 'Status *' })",
        STATUS,
    )

    # --------------------------------------------------------
    # Set Job Start
    # --------------------------------------------------------

    print()
    print(f"Setting start time: {START_TIME}")

    pw(
        "fill",
        "getByRole('textbox').nth(2)",
        START_TIME,
    )

    # --------------------------------------------------------
    # Set Job End
    # --------------------------------------------------------

    print()
    print(f"Setting end time: {END_TIME}")

    pw(
        "fill",
        "getByRole('textbox').nth(4)",
        END_TIME,
    )

    # --------------------------------------------------------
    # Set Charge Type
    # --------------------------------------------------------

    print()
    print(f"Setting charge type: {CHARGE_TYPE}")

    pw(
        "select",
        "getByRole('combobox', { name: 'Charge Type *' })",
        CHARGE_TYPE,
    )

    # --------------------------------------------------------
    # Final snapshot before saving
    # --------------------------------------------------------

    print()
    print("Final snapshot before Save...")

    pw("snapshot")

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    print()
    print("Saving Worklog...")

    pw(
        "click",
        "getByRole('button', { name: 'Save' })",
    )

    # --------------------------------------------------------
    # Verify resulting page
    # --------------------------------------------------------

    time.sleep(1)

    print()
    print("Final page state:")

    pw("snapshot")

    print()
    print("Worklog automation complete.")


# ============================================================
# Entry point
# ============================================================

if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        print("\nCancelled.")

    except Exception as e:
        print()
        print("ERROR:")
        print(e)
        sys.exit(1)