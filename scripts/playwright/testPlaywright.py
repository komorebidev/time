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
        "Could not find playwright-cli."
    )


PLAYWRIGHT_CLI = find_playwright_cli()


# ============================================================
# Run playwright-cli
# ============================================================

def pw(*args, check=True):
    """
    Run playwright-cli directly.

    subprocess.run() passes each argument separately, so URLs
    containing '&' do not need shell escaping.
    """

    command = [
        PLAYWRIGHT_CLI,
        f"--s={SESSION}",
        *args,
    ]

    print(
        ">",
        " ".join(
            f'"{arg}"' if any(c in arg for c in " &")
            else arg
            for arg in command
        )
    )

    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        shell=False,
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
# Attach to Microsoft Edge
# ============================================================

def attach():
    print("Attaching to Microsoft Edge...")

    command = [
        PLAYWRIGHT_CLI,
        "attach",
        "--extension=msedge",
        f"--session={SESSION}",
    ]

    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        shell=False,
    )

    if result.stdout:
        print(result.stdout)

    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        raise RuntimeError(
            "Could not attach to Microsoft Edge.\n\n"
            + result.stdout
            + "\n"
            + result.stderr
        )

    time.sleep(1)


# ============================================================
# Main
# ============================================================

def main():

    print()
    print(f"Using playwright-cli: {PLAYWRIGHT_CLI}")
    print()

    # --------------------------------------------------------
    # Attach to existing Edge
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
    # Open ticket
    # --------------------------------------------------------

    print()
    print(f"Opening ticket {ticket}...")

    pw("goto", url)

    # --------------------------------------------------------
    # Snapshot
    # --------------------------------------------------------

    print()
    print("Taking ticket snapshot...")

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

    time.sleep(1)

    # --------------------------------------------------------
    # Snapshot Worklog
    # --------------------------------------------------------

    print()
    print("Taking Worklog snapshot...")

    pw("snapshot")

    # --------------------------------------------------------
    # Enter Worklog
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
    # Set Start Time
    # --------------------------------------------------------

    print()
    print(f"Setting start time: {START_TIME}")

    pw(
        "fill",
        "getByRole('textbox').nth(2)",
        START_TIME,
    )

    # --------------------------------------------------------
    # Set End Time
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
    # Final snapshot
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

    time.sleep(1)

    # --------------------------------------------------------
    # Verify
    # --------------------------------------------------------

    print()
    print("Verifying result...")

    pw("snapshot")

    print()
    print("========================================")
    print("Worklog automation complete.")
    print("========================================")


# ============================================================
# Entry point
# ============================================================

if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        print()
        print("Cancelled.")

    except Exception as e:
        print()
        print("ERROR:")
        print(e)
        sys.exit(1)