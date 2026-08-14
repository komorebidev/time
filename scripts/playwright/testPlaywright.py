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
        "Could not find playwright-cli.\n"
        "Run:\n"
        "python -c \"import shutil; print(shutil.which('playwright-cli'))\""
    )


PLAYWRIGHT_CLI = find_playwright_cli()


# ============================================================
# Run playwright-cli
# ============================================================

def pw(*args, check=True):
    """
    Run playwright-cli through cmd.exe.

    This is important on Windows because playwright-cli is
    installed as a .cmd file and URLs contain '&'.
    """

    formatted_args = []

    for arg in args:
        # Quote arguments containing spaces or shell-sensitive
        # characters.
        if any(c in arg for c in ' &'):
            escaped = arg.replace('"', '\\"')
            formatted_args.append(f'"{escaped}"')
        else:
            formatted_args.append(arg)

    command = (
        f'"{PLAYWRIGHT_CLI}" '
        f'--s={SESSION} '
        f'{" ".join(formatted_args)}'
    )

    print(f"> {command}")

    result = subprocess.run(
        ["cmd.exe", "/d", "/c", command],
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
            f"{command}\n\n"
            f"Exit code: {result.returncode}"
        )

    return result


# ============================================================
# Attach to Microsoft Edge
# ============================================================

def attach():
    print("Attaching to Microsoft Edge...")

    command = (
        f'"{PLAYWRIGHT_CLI}" '
        f'attach '
        f'--extension=msedge '
        f'--session={SESSION}'
    )

    result = subprocess.run(
        ["cmd.exe", "/d", "/c", command],
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
            f"{result.stdout}\n"
            f"{result.stderr}"
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
    # Attach to Edge
    # --------------------------------------------------------

    attach()

    # --------------------------------------------------------
    # Ticket number
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
    # Snapshot ticket
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
    # Snapshot Worklog form
    # --------------------------------------------------------

    print()
    print("Taking Worklog snapshot...")

    pw("snapshot")

    # --------------------------------------------------------
    # Worklog text
    # --------------------------------------------------------

    print()
    print("Entering worklog...")

    pw(
        "fill",
        "locator('[contenteditable=\"true\"]').first()",
        WORKLOG_TEXT,
    )

    # --------------------------------------------------------
    # Status
    # --------------------------------------------------------

    print()
    print(f"Setting status: {STATUS}")

    pw(
        "select",
        "getByRole('combobox', { name: 'Status *' })",
        STATUS,
    )

    # --------------------------------------------------------
    # Start time
    # --------------------------------------------------------

    print()
    print(f"Setting start time: {START_TIME}")

    pw(
        "fill",
        "getByRole('textbox').nth(2)",
        START_TIME,
    )

    # --------------------------------------------------------
    # End time
    # --------------------------------------------------------

    print()
    print(f"Setting end time: {END_TIME}")

    pw(
        "fill",
        "getByRole('textbox').nth(4)",
        END_TIME,
    )

    # --------------------------------------------------------
    # Charge type
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