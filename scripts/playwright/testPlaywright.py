import os
import sys
import atexit
import platform
import re
import shutil
import stat
import subprocess
import tempfile
import textwrap
import time


BASE_URL = "https://support.eiresystems.com/ticket"

ASSIGNED_TICKETS_URL = (
    "https://support.eiresystems.com/tickets"
    "?area=1"
    "&mainview=team"
    "&viewid=2"
    "&selid=75"
    "&sellevel=2"
    "&selparentid=engineers%20tky"
)

SESSION = "halo"


# ============================================================
# PLAYWRIGHT CLI
# ============================================================

def find_playwright_cli():
    """
    Find playwright-cli without hard-coding the Windows username.
    """
    cli = shutil.which("playwright-cli")

    if cli:
        return cli

    cli = shutil.which("playwright-cli.cmd")

    if cli:
        return cli

    raise FileNotFoundError(
        "playwright-cli was not found in PATH."
    )


CLI = find_playwright_cli()


def run_cli(*args, check=True):
    """
    Run playwright-cli safely across platforms.
    """

    if platform.system() == "Windows":

        quoted_args = []

        for arg in args:

            s = str(arg)

            if not (
                s.startswith('"')
                and s.endswith('"')
            ):
                s = f'"{s}"'

            quoted_args.append(s)

        cmd_str = (
            f'"{CLI}" "--s={SESSION}" '
            + " ".join(quoted_args)
        )

        print()
        print("> " + cmd_str)

        result = subprocess.run(
            cmd_str,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            shell=True,
        )

    else:

        command = [
            CLI,
            f"--s={SESSION}",
            *[str(arg) for arg in args]
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
            shell=False,
        )

    if result.stdout:
        print(result.stdout)

    if (
        check
        and result.returncode != 0
    ):
        raise RuntimeError(
            "Playwright CLI command failed "
            f"with exit code {result.returncode}"
        )

    return result.stdout


# ============================================================
# CLEANUP
# ============================================================

def remove_readonly(func, path, excinfo):
    """
    Error handler for shutil.rmtree on Windows.
    """
    os.chmod(
        path,
        stat.S_IWRITE
    )

    func(path)


def cleanup_local_artifacts():
    """
    Remove local Playwright/session folders.
    """

    folders_to_remove = [
        ".playwright",
        ".playwright-cli"
    ]

    for folder in folders_to_remove:

        if (
            os.path.exists(folder)
            and os.path.isdir(folder)
        ):

            for attempt in range(3):

                try:

                    shutil.rmtree(
                        folder,
                        onerror=remove_readonly
                    )

                    print(
                        f"Cleaned up local folder: {folder}"
                    )

                    break

                except Exception as e:

                    if attempt == 2:

                        print(
                            f"Note: Could not fully remove "
                            f"folder {folder}: {e}"
                        )

                    else:

                        time.sleep(0.5)


atexit.register(
    cleanup_local_artifacts
)


# ============================================================
# INPUT VALIDATION
# ============================================================

def normalize_time(value, field_name):
    """
    Validate and normalize HH:MM input.

    Accepts:
        08:00
        8:00

    Returns:
        08:00

    Empty input means leave unchanged.
    """

    value = value.strip()

    if not value:
        return ""

    match = re.fullmatch(
        r"(\d{1,2}):(\d{2})",
        value
    )

    if not match:
        raise ValueError(
            f"{field_name} must be HH:MM, e.g. 08:00"
        )

    hour = int(match.group(1))
    minute = int(match.group(2))

    if hour < 0 or hour > 23:
        raise ValueError(
            f"{field_name} hour must be between 00 and 23."
        )

    if minute < 0 or minute > 59:
        raise ValueError(
            f"{field_name} minute must be between 00 and 59."
        )

    return f"{hour:02d}:{minute:02d}"


# ============================================================
# EDGE / NAVIGATION
# ============================================================

def attach():
    """
    Attach playwright-cli to the existing Microsoft Edge session.
    """

    print(
        "\nAttaching to Microsoft Edge..."
    )

    run_cli(
        "attach",
        "--extension=msedge",
        check=True
    )

    time.sleep(1)

    print(
        "\nOpening Assigned Tickets view..."
    )

    run_cli(
        "goto",
        ASSIGNED_TICKETS_URL,
        check=True
    )

    time.sleep(2)


def goto_ticket(ticket):
    """
    Navigate directly to the requested Halo ticket.
    """

    url = (
        f"{BASE_URL}"
        f"?id={ticket}"
        f"&showalltickettypes=1"
    )

    print(
        f"\nOpening ticket {ticket}..."
    )

    run_cli(
        "goto",
        url,
        check=True
    )

    time.sleep(1)


# ============================================================
# SNAPSHOT / TICKET PARSING
# ============================================================

def parse_snapshot_tickets(snapshot_output):
    """
    Parse Playwright CLI snapshot output
    into ticket dictionaries.
    """

    tickets = []

    lines = snapshot_output.splitlines()

    current_block = []
    in_ticket_block = False

    for line in lines:

        if (
            '"Bulk select"' in line
            or '[cursor=pointer]' in line
        ):

            if current_block:

                parsed = _parse_single_block(
                    current_block
                )

                if parsed:
                    tickets.append(parsed)

            current_block = [line]
            in_ticket_block = True

        elif in_ticket_block:

            current_block.append(line)

    if current_block:

        parsed = _parse_single_block(
            current_block
        )

        if parsed:
            tickets.append(parsed)

    return tickets


def _parse_single_block(block_lines):

    ticket_id = None
    company = None
    status = None
    ticket_name = None
    date_str = None
    ticket_type = None
    total_hours = None

    clean_lines = []

    for line in block_lines:

        match_quote = re.search(
            r'"([^"]+)"',
            line
        )

        match_text = re.search(
            r'text:\s*(.*)',
            line
        )

        if match_quote:

            clean_lines.append(
                match_quote.group(1)
            )

        elif match_text:

            clean_lines.append(
                match_text.group(1).strip()
            )

        else:

            parts = line.split(
                ':',
                1
            )

            if len(parts) > 1:

                if '"' not in parts[1]:

                    val = parts[1].strip()

                    if (
                        val
                        and not val.startswith('[')
                    ):

                        clean_lines.append(val)

    for i, l in enumerate(clean_lines):

        if re.match(
            r'^00\d{5}$',
            l
        ):

            ticket_id = l

        elif (
            '/' in l
            and 'EIRE' in l
        ):

            company = l

        elif l in [
            "Completed (On Hold)",
            "In Progress",
            "On Hold",
            "New",
            "Closed"
        ]:

            status = l

        elif re.match(
            r'^\d{1,2}/\d{1,2}/\d{4}\s+\d{2}:\d{2}$',
            l
        ):

            date_str = l

            if i > 0:

                ticket_name = clean_lines[
                    i - 1
                ]

        elif l in [
            "Service Request",
            "Incident",
            "Project Support",
            "Problem",
            "Change Request"
        ]:

            ticket_type = l

        elif re.match(
            r'^\d+:\d{2}$',
            l
        ):

            total_hours = l

    if ticket_id:

        return {
            "id": ticket_id,
            "title": (
                ticket_name
                or f"Ticket {ticket_id}"
            ),
            "company": company,
            "status": status,
            "date": date_str,
            "type": ticket_type,
            "hours": total_hours
        }

    return None


def scrape_ticket_options():

    print(
        "\nTaking snapshot to parse tickets..."
    )

    output = run_cli(
        "snapshot",
        check=True
    )

    return parse_snapshot_tickets(
        output
    )


# ============================================================
# HALO WORKLOG AUTOMATION
# ============================================================

def run_halo_automation(
    worklog_text,
    status,
    start_time,
    end_time,
    charge_type
):
    """
    Run HaloPSA Worklog automation.

    IMPORTANT ORDER:

    1. Open Worklog
    2. Find Job Start
    3. Click Job Start
    4. Fill Job Start
    5. Verify Job Start
    6. Continue with remaining fields

    Job Start:
        input[name="actionarrivaldate_time"]

    Job End:
        input[name="actioncompletiondate_time"]

    We deliberately DO NOT identify these fields
    by their position among all inputs.
    """

    js_code = textwrap.dedent(
        f"""
        async page => {{

            console.log("");
            console.log("========================================");
            console.log("HALO WORKLOG AUTOMATION");
            console.log("========================================");


            // ====================================================
            // 1. OPEN WORKLOG
            // ====================================================

            console.log("");
            console.log("STEP 1: Opening Worklog...");

            const worklogButton = page.getByRole(
                "button",
                {{
                    name: "Worklog"
                }}
            );

            await worklogButton.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            await worklogButton.click();

            await page.waitForTimeout(1500);


            // ====================================================
            // 2. FIND JOB START
            // ====================================================

            console.log("");
            console.log("STEP 2: Finding Job Start...");

            const startTimeInput = page.locator(
                'input[name="actionarrivaldate_time"]'
            );

            const startCount =
                await startTimeInput.count();

            console.log(
                "Job Start matching elements:",
                startCount
            );

            if (startCount !== 1) {{
                throw new Error(
                    "Expected exactly one Job Start input, found "
                    + startCount
                );
            }}

            await startTimeInput.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            console.log(
                "Job Start found."
            );

            console.log(
                "Job Start current value:",
                await startTimeInput.inputValue()
            );


            // ====================================================
            // 3. CLICK JOB START
            // ====================================================

            console.log("");
            console.log("STEP 3: Clicking Job Start...");

            await startTimeInput.click();

            await page.waitForTimeout(150);


            const focusedAfterClick =
                await startTimeInput.evaluate(
                    el => document.activeElement === el
                );

            console.log(
                "Job Start focused after click:",
                focusedAfterClick
            );

            if (!focusedAfterClick) {{
                throw new Error(
                    "Job Start was located and clicked, "
                    + "but it did not receive focus."
                );
            }}


            // ====================================================
            // 4. FILL JOB START
            // ====================================================

            if ({start_time!r}) {{

                console.log("");
                console.log("STEP 4: Filling Job Start...");

                console.log(
                    "Requested Job Start:",
                    {start_time!r}
                );

                await startTimeInput.fill(
                    {start_time!r}
                );

                await page.waitForTimeout(150);


                const startAfterFill =
                    await startTimeInput.inputValue();

                console.log(
                    "Job Start after fill:",
                    startAfterFill
                );


                if (
                    startAfterFill !== {start_time!r}
                ) {{

                    throw new Error(
                        "Job Start did not contain the "
                        + "requested value after fill. "
                        + "Expected: "
                        + {start_time!r}
                        + " Actual: "
                        + startAfterFill
                    );
                }}

            }} else {{

                console.log(
                    "No Job Start supplied; "
                    + "leaving existing value unchanged."
                );

            }}


            // ====================================================
            // JOB END LOCATOR
            // ====================================================

            console.log("");
            console.log(
                "Finding Job End..."
            );

            const endTimeInput = page.locator(
                'input[name="actioncompletiondate_time"]'
            );

            await endTimeInput.waitFor({{
                state: "visible",
                timeout: 10000
            }});


            // ====================================================
            // WORKLOG TEXT
            // ====================================================

            console.log("");
            console.log(
                "Entering worklog text..."
            );

            const editor = page.locator(
                '[contenteditable="true"]'
            ).first();

            await editor.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            await editor.fill(
                {worklog_text!r}
            );

            await page.waitForTimeout(300);


            // ====================================================
            // STATUS
            // ====================================================

            console.log("");
            console.log(
                "Setting status: {status}"
            );

            const statusCombobox = page.getByRole(
                "combobox",
                {{
                    name: "Status *"
                }}
            );

            await statusCombobox.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            await statusCombobox.click();

            await page.waitForTimeout(500);


            const statusOption = page.locator(
                ".Select__option",
                {{
                    hasText: {status!r}
                }}
            ).last();


            await statusOption.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            await statusOption.click();

            await page.waitForTimeout(500);


            // ====================================================
            // JOB END
            // ====================================================

            if ({end_time!r}) {{

                console.log("");
                console.log("Setting Job End...");

                console.log(
                    "Requested Job End:",
                    {end_time!r}
                );

                await endTimeInput.click();

                await endTimeInput.fill(
                    {end_time!r}
                );

                await page.waitForTimeout(150);


                const endAfterFill =
                    await endTimeInput.inputValue();

                console.log(
                    "Job End after fill:",
                    endAfterFill
                );


                if (
                    endAfterFill !== {end_time!r}
                ) {{

                    throw new Error(
                        "Job End did not contain the "
                        + "requested value after fill. "
                        + "Expected: "
                        + {end_time!r}
                        + " Actual: "
                        + endAfterFill
                    );
                }}

            }} else {{

                console.log(
                    "No Job End supplied; "
                    + "leaving existing value unchanged."
                );

            }}


            // ====================================================
            // VERIFY BOTH TIMES
            // ====================================================

            console.log("");
            console.log("========================================");
            console.log("TIME VERIFICATION");
            console.log("========================================");

            const verifiedStart =
                await startTimeInput.inputValue();

            const verifiedEnd =
                await endTimeInput.inputValue();

            console.log(
                "Job Start:",
                verifiedStart
            );

            console.log(
                "Job End:",
                verifiedEnd
            );


            if (
                {start_time!r}
                && verifiedStart !== {start_time!r}
            ) {{

                throw new Error(
                    "FINAL Job Start verification failed."
                );
            }}

            if (
                {end_time!r}
                && verifiedEnd !== {end_time!r}
            ) {{

                throw new Error(
                    "FINAL Job End verification failed."
                );
            }}


            // ====================================================
            // CHARGE TYPE
            // ====================================================

            console.log("");
            console.log(
                "Setting charge type: {charge_type}"
            );

            const chargeCombobox = page.getByRole(
                "combobox",
                {{
                    name: "Charge Type *"
                }}
            );

            await chargeCombobox.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            await chargeCombobox.click();

            await page.waitForTimeout(500);


            const chargeOption = page.locator(
                ".Select__option",
                {{
                    hasText: {charge_type!r}
                }}
            ).last();


            await chargeOption.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            await chargeOption.click();

            await page.waitForTimeout(500);


            // ====================================================
            // FINAL CHECK BEFORE SAVE
            // ====================================================

            console.log("");
            console.log("========================================");
            console.log("FINAL VALUES BEFORE SAVE");
            console.log("========================================");

            const finalStart =
                await startTimeInput.inputValue();

            const finalEnd =
                await endTimeInput.inputValue();

            console.log(
                "Job Start:",
                finalStart
            );

            console.log(
                "Job End:",
                finalEnd
            );


            if (
                {start_time!r}
                && finalStart !== {start_time!r}
            ) {{

                throw new Error(
                    "Job Start changed unexpectedly before save."
                    + " Expected "
                    + {start_time!r}
                    + " but found "
                    + finalStart
                );
            }}

            if (
                {end_time!r}
                && finalEnd !== {end_time!r}
            ) {{

                throw new Error(
                    "Job End changed unexpectedly before save."
                    + " Expected "
                    + {end_time!r}
                    + " but found "
                    + finalEnd
                );
            }}


            // ====================================================
            // SAVE
            // ====================================================

            console.log("");
            console.log(
                "Saving worklog..."
            );

            const saveButton = page.getByRole(
                "button",
                {{
                    name: "Save",
                    exact: true
                }}
            );

            await saveButton.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            await saveButton.click();

            await page.waitForTimeout(2000);


            console.log("");
            console.log("========================================");
            console.log("WORKLOG SAVED");
            console.log("========================================");

            console.log(
                "URL:",
                page.url()
            );


            return {{
                finalStart,
                finalEnd,
                urlAfterSave: page.url()
            }};
        }}
        """
    )


    temp_path = None

    try:

        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".js",
            prefix="halo_worklog_",
            delete=False,
            encoding="utf-8"
        ) as temp_file:

            temp_file.write(
                js_code
            )

            temp_path = temp_file.name


        print(
            "\nRunning Halo automation..."
        )


        run_cli(
            "run-code",
            f"--filename={temp_path}",
            check=True
        )


    finally:

        if (
            temp_path
            and os.path.exists(temp_path)
        ):

            try:

                os.remove(
                    temp_path
                )

            except OSError:

                pass


# ============================================================
# FINAL SNAPSHOT
# ============================================================

def take_snapshot():

    print(
        "\nTaking final snapshot..."
    )

    run_cli(
        "snapshot",
        check=True
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print(
        "HaloPSA Worklog Automation"
    )
    print(
        "=========================="
    )


    try:

        # ====================================================
        # ATTACH
        # ====================================================

        attach()


        # ====================================================
        # SELECT TICKET
        # ====================================================

        ticket = ""

        choice = input(
            "\n[1] Enter Ticket ID manually\n"
            "[2] Scrape Ticket IDs & metadata "
            "from current snapshot view\n"
            "Select option [1/2]: "
        ).strip()


        if choice == "2":

            tickets = scrape_ticket_options()

            if tickets:

                print(
                    f"\nFound {len(tickets)} "
                    "tickets on the current page:"
                )


                for idx, t in enumerate(
                    tickets,
                    1
                ):

                    date_info = (
                        f" [{t['date']}]"
                        if t["date"]
                        else ""
                    )

                    type_info = (
                        f" ({t['type']})"
                        if t["type"]
                        else ""
                    )

                    print(
                        f"  [{idx}] "
                        f"{t['id']}"
                        f"{date_info}"
                        f"{type_info}"
                        f" - {t['title']}"
                    )


                sel = input(
                    "\nEnter selection number "
                    "or type a Ticket ID directly: "
                ).strip()


                if (
                    sel.isdigit()
                    and 1 <= int(sel) <= len(tickets)
                ):

                    ticket = tickets[
                        int(sel) - 1
                    ]["id"]

                    print(
                        f"Selected Ticket ID: {ticket}"
                    )

                else:

                    ticket = sel

            else:

                print(
                    "No tickets could be automatically "
                    "parsed from this snapshot."
                )


        # ====================================================
        # VALIDATE TICKET
        # ====================================================

        while (
            not ticket
            or not ticket.isdigit()
        ):

            ticket = input(
                "\nEnter Ticket Number: "
            ).strip()


            if (
                not ticket
                or not ticket.isdigit()
            ):

                print(
                    "Please enter a valid "
                    "numeric ticket number."
                )

                ticket = ""


        # ====================================================
        # WORKLOG TEXT
        # ====================================================

        worklog_text = ""

        while not worklog_text:

            worklog_text = input(
                "Worklog text (Required): "
            ).strip()


            if not worklog_text:

                print(
                    "Worklog text cannot be empty."
                )


        # ====================================================
        # STATUS
        # ====================================================

        status_options = [
            "In Progress",
            "Completed (On Hold)",
            "On Hold"
        ]

        default_status = (
            "Completed (On Hold)"
        )

        status = default_status


        while True:

            status_input = input(
                f"Status [? for options] "
                f"[{default_status}]: "
            ).strip()


            if status_input == "?":

                print(
                    "\nAvailable Statuses:"
                )

                print(
                    "-" * 25
                )


                for idx, opt in enumerate(
                    status_options,
                    1
                ):

                    print(
                        f"  [{idx}] {opt}"
                    )


                print(
                    "-" * 25
                )


                sel = input(
                    "Select option number: "
                ).strip()


                if (
                    sel.isdigit()
                    and 1 <= int(sel) <= len(status_options)
                ):

                    status = status_options[
                        int(sel) - 1
                    ]

                    break

                else:

                    print(
                        "Invalid selection. "
                        "Try again."
                    )


            elif not status_input:

                status = default_status

                break


            else:

                status = status_input

                break


        # ====================================================
        # TIMES
        # ====================================================

        while True:

            try:

                start_time = normalize_time(
                    input(
                        "Start time "
                        "[Leave unchanged, e.g. 09:00]: "
                    ),
                    "Start time"
                )

                break

            except ValueError as exc:

                print(
                    f"Error: {exc}"
                )


        while True:

            try:

                end_time = normalize_time(
                    input(
                        "End time "
                        "[Leave unchanged, e.g. 10:00]: "
                    ),
                    "End time"
                )

                break

            except ValueError as exc:

                print(
                    f"Error: {exc}"
                )


        # ====================================================
        # CHARGE TYPE
        # ====================================================

        charge_options = [
            "Project Work- Managed Services",
            "Research (work-specific)",
            "Professional Development",
            "Internal Work"
        ]

        default_charge = "Internal Work"

        charge_type = default_charge


        while True:

            charge_input = input(
                f"Charge type [? for options] "
                f"[{default_charge}]: "
            ).strip()


            if charge_input == "?":

                print(
                    "\nAvailable Charge Types:"
                )

                print(
                    "-" * 35
                )


                for idx, opt in enumerate(
                    charge_options,
                    1
                ):

                    print(
                        f"  [{idx}] {opt}"
                    )


                print(
                    "-" * 35
                )


                sel = input(
                    "Select option number: "
                ).strip()


                if (
                    sel.isdigit()
                    and 1 <= int(sel) <= len(charge_options)
                ):

                    charge_type = charge_options[
                        int(sel) - 1
                    ]

                    break

                else:

                    print(
                        "Invalid selection. "
                        "Try again."
                    )


            elif not charge_input:

                charge_type = default_charge

                break


            else:

                charge_type = charge_input

                break


        # ====================================================
        # OPEN TICKET
        # ====================================================

        goto_ticket(
            ticket
        )


        # ====================================================
        # RUN AUTOMATION
        # ====================================================

        run_halo_automation(
            worklog_text,
            status,
            start_time,
            end_time,
            charge_type
        )


        # ====================================================
        # FINAL SNAPSHOT
        # ====================================================

        take_snapshot()


        print()
        print(
            "================================"
        )
        print(
            "Worklog automation completed."
        )
        print(
            "================================"
        )


    except Exception as exc:

        print()
        print(
            "ERROR:"
        )

        print(
            exc
        )

        sys.exit(1)


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    main()