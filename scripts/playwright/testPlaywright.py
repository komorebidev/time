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
    sys.exit(1)


# --------------------------------------------------
# Run Playwright CLI
# --------------------------------------------------

def run_cli(args):
    """
    Run playwright-cli safely on Windows.

    Handles:
      - npm's .cmd launcher
      - URLs containing &
      - UTF-8 Playwright output
    """

    args = [str(arg) for arg in args]

    if (
        sys.platform == "win32"
        and PLAYWRIGHT_CLI.lower().endswith(".cmd")
    ):
        quoted_args = []

        for arg in args:
            escaped = arg.replace('"', '\\"')
            quoted_args.append(f'"{escaped}"')

        command = (
            f'"{PLAYWRIGHT_CLI}" '
            + " ".join(quoted_args)
        )

        result = subprocess.run(
            command,
            shell=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
        )

    else:
        result = subprocess.run(
            [PLAYWRIGHT_CLI, *args],
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
        )

    if result.stdout:
        print(result.stdout)

    if result.returncode != 0:
        if result.stderr:
            print(result.stderr, file=sys.stderr)

        raise RuntimeError(
            "Playwright CLI command failed:\n"
            + " ".join(args)
        )

    return result.stdout


# --------------------------------------------------
# Playwright helper
# --------------------------------------------------

def pw(*args):
    return run_cli([
        f"-s={SESSION}",
        *args,
    ])


# --------------------------------------------------
# Attach to Microsoft Edge
# --------------------------------------------------

def attach():

    print("Attaching to Microsoft Edge...")

    run_cli([
        "attach",
        "--extension=msedge",
        f"-s={SESSION}",
    ])


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
    print("Taking snapshot...")

    pw("snapshot")

    print()
    print("Done.")


# --------------------------------------------------
# Entry point
# --------------------------------------------------

if __name__ == "__main__":
    main()