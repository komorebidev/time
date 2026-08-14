import shutil
import subprocess
import sys
import time


BASE_URL = "https://support.eiresystems.com/ticket"

SESSION = "halo"

WORKLOG_TEXT = (
    "Eire: Email catchup, internal communication, time recording, work logs"
)

STATUS = "Completed (On Hold)"
CHARGE_TYPE = "Internal work"

# Hard-coded for this version.
START_TIME = "09:00"
END_TIME = "10:00"


def find_cli():
    """Find playwright-cli without hard-coding the Windows username."""

    cli = shutil.which("playwright-cli")

    if cli:
        return cli

    cli = shutil.which("playwright-cli.cmd")

    if cli:
        return cli

    raise FileNotFoundError(
        "playwright-cli was not found in PATH."
    )


CLI = find_cli()


def pw(*args, check=True):
    """
    Run:

        playwright-cli --s=halo <command> ...

    directly.

    Using a list of arguments is important because the ticket URL
    contains '&'.
    """

    command = [
        CLI,
        f"--s={SESSION}",
        *[str(arg) for arg in args],
    ]

    print()
    print("> " + " ".join(command))

    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    if result.stdout:
        print(result.stdout)

    if check and result.returncode != 0:
        raise RuntimeError(
            f"Playwright CLI command failed:\n"
            f"{' '.join(command)}\n"
            f"Exit code: {result.returncode}"
        )

    return result.stdout


def attach():
    print("Attaching to Microsoft Edge...")

    pw("attach")

    time.sleep(1)


def goto_ticket(ticket):
    url = (
        f"{BASE_URL}"
        f"?id={ticket}"
        f"&showalltickettypes=1"
    )

    print(f"\nOpening ticket {ticket}...")

    pw("goto", url)

    time.sleep(1)


def open_worklog():
    print("\nOpening Worklog...")

    pw("click", "button", "Worklog")

    time.sleep(1)


def enter_worklog():
    print("\nEntering worklog...")

    pw(
        "fill",
        '[contenteditable="true"]',
        WORKLOG_TEXT,
    )


def set_status():
    print(f"\nSetting status: {STATUS}")

    pw(
        "click",
        "combobox",
        "Status *",
    )

    time.sleep(0.5)

    pw(
        "click",
        "text",
        STATUS,
    )


def set_times():
    """
    Set the two time fields.

    From the snapshot, the Worklog contains:

        Job Start
            Date textbox
            time textbox

        Job End
            Date textbox
            time textbox

    The date fields have the accessible name "Date", while the
    time fields do not.

    Therefore we locate the text inputs and use the two time
    inputs associated with the Worklog form.
    """

    print(f"\nSetting Job Start: {START_TIME}")
    print(f"Setting Job End:   {END_TIME}")

    # First try the two time inputs using their current values.
    #
    # This is intentionally done with JavaScript through the CLI
    # rather than guessing at generated ref IDs.

    script = f"""
    const inputs = [...document.querySelectorAll('input')];

    const timeInputs = inputs.filter(input => {{
        const value = input.value || '';
        const type = (input.type || '').toLowerCase();

        return (
            type === 'text' &&
            /^\\d{{1,2}}:\\d{{2}}$/.test(value)
        );
    }});

    if (timeInputs.length < 2) {{
        throw new Error(
            'Could not find the two Worklog time inputs. Found: '
            + timeInputs.length
        );
    }}

    timeInputs[0].value = '{START_TIME}';
    timeInputs[0].dispatchEvent(
        new Event('input', {{ bubbles: true }})
    );
    timeInputs[0].dispatchEvent(
        new Event('change', {{ bubbles: true }})
    );

    timeInputs[1].value = '{END_TIME}';
    timeInputs[1].dispatchEvent(
        new Event('input', {{ bubbles: true }})
    );
    timeInputs[1].dispatchEvent(
        new Event('change', {{ bubbles: true }})
    );
    """

    pw("eval", script)


def set_charge_type():
    print(f"\nSetting charge type: {CHARGE_TYPE}")

    pw(
        "click",
        "combobox",
        "Charge Type *",
    )

    time.sleep(0.5)

    pw(
        "click",
        "text",
        CHARGE_TYPE,
    )


def take_snapshot():
    print("\nTaking final snapshot...")

    pw("snapshot")


def save():
    print("\nSaving worklog...")

    pw(
        "click",
        "button",
        "Save",
    )

    time.sleep(2)


def main():
    print("HaloPSA Worklog Automation")
    print("==========================")

    ticket = input("Ticket number: ").strip()

    if not ticket.isdigit():
        print("Invalid ticket number.")
        return

    try:
        attach()

        goto_ticket(ticket)

        print("\nTicket opened.")

        open_worklog()

        enter_worklog()

        set_status()

        set_times()

        set_charge_type()

        take_snapshot()

        save()

        print("\nDone.")

    except Exception as exc:
        print("\nERROR:")
        print(exc)
        sys.exit(1)


if __name__ == "__main__":
    main()