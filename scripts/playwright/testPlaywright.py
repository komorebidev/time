import shutil
import subprocess
import sys


# --------------------------------------------------
# Configuration
# --------------------------------------------------

SESSION = "halo"

BASE_URL = "https://support.eiresystems.com/ticket"

WORKLOG_TEXT = (
    "Eire: Email catchup, internal communication, "
    "time recording, work logs"
)

STATUS = "Completed (On Hold)"
CHARGE_TYPE = "Internal work"

START_TIME = "09:00"
END_TIME = "10:00"


# --------------------------------------------------
# Locate Playwright CLI
# --------------------------------------------------

PLAYWRIGHT_CLI = shutil.which("playwright-cli")

if not PLAYWRIGHT_CLI:
    print(
        "ERROR: playwright-cli was not found on PATH.",
        file=sys.stderr,
    )
    print(
        "Run: python -c "
        "\"import shutil; print(shutil.which('playwright-cli'))\"",
        file=sys.stderr,
    )
    sys.exit(1)


# --------------------------------------------------
# Playwright CLI helper
# --------------------------------------------------

def pw(*args):
    """Run a Playwright CLI command against our named session."""

    command = [
        PLAYWRIGHT_CLI,
        f"-s={SESSION}",
        *args,
    ]

    result = subprocess.run(
        command,
        text=True,
        capture_output=True,
    )

    if result.stdout:
        print(result.stdout)

    if result.returncode != 0:
        if result.stderr:
            print(result.stderr, file=sys.stderr)

        raise RuntimeError(
            f"Playwright CLI command failed:\n"
            f"{' '.join(command)}"
        )

    return result.stdout


# --------------------------------------------------
# Attach to Microsoft Edge
# --------------------------------------------------

def attach():
    print("Attaching to Microsoft Edge...")

    command = [
        PLAYWRIGHT_CLI,
        "attach",
        "--extension=msedge",
        f"-s={SESSION}",
    ]

    result = subprocess.run(
        command,
        text=True,
        capture_output=True,
    )

    if result.stdout:
        print(result.stdout)

    if result.returncode != 0:
        if result.stderr:
            print(result.stderr, file=sys.stderr)

        raise RuntimeError(
            "Could not attach to Microsoft Edge."
        )


# --------------------------------------------------
# Main
# --------------------------------------------------

def main():

    # Attach to the existing Edge session.
    attach()

    # Ask for ticket number.
    ticket = input("Ticket number: ").strip()

    if not ticket.isdigit():
        print("Invalid ticket number.")
        return

    # Build ticket URL.
    url = (
        f"{BASE_URL}"
        f"?id={ticket}"
        f"&showalltickettypes=1"
    )

    print()
    print(f"Opening ticket {ticket}...")

    pw("goto", url)

    print()
    print("Taking page snapshot...")

    pw("snapshot")

    print()
    print("Done.")


# --------------------------------------------------
# Entry point
# --------------------------------------------------

if __name__ == "__main__":
    main()