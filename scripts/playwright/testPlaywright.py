import os
import sys
import atexit
import platform
import re
import shutil
import stat
import subprocess
import tempfile
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
    Error handler for shutil.rmtree to clear read-only bits
    and retry on Windows.
    """

    os.chmod(
        path,
        stat.S_IWRITE
    )

    func(path)


def cleanup_local_artifacts():
    """
    Remove local Playwright folders created in the working
    directory.
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
# EDGE / NAVIGATION
# ============================================================

def attach():
    """
    Attach to the existing Microsoft Edge session using
    the Playwright CLI extension.
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
    Navigate directly to a Halo ticket.
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
# TICKET SNAPSHOT PARSING
# ============================================================

def parse_snapshot_tickets(snapshot_output):
    """
    Parse playwright-cli snapshot text into ticket dictionaries.
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

            if (
                len(parts) > 1
                and '"' not in parts[1]
            ):

                val = parts[1].strip()

                if (
                    val
                    and not val.startswith('[')
                ):

                    clean_lines.append(
                        val
                    )

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

                ticket_name = (
                    clean_lines[i - 1]
                )

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

    The Job Start and Job End inputs are selected by their
    stable HTML name attributes rather than snapshot refs,
    generated IDs, or input indexes.

    Known fields:

        Job Start:
        input[name="actionarrivaldate_time"]

        Job End:
        input[name="actioncompletiondate_time"]
    """

    js_code = f"""
async page => {{

    // ========================================================
    // OPEN WORKLOG
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("OPENING WORKLOG");
    console.log("========================================");

    await page.getByRole(
        "button",
        {{ name: "Worklog" }}
    ).click();

    await page.waitForTimeout(1500);


    // ========================================================
    // WORKLOG TEXT
    // ========================================================

    console.log(
        "Entering worklog..."
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


    // ========================================================
    // EXACT TIME INPUTS
    // ========================================================

    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );

    await start.waitFor({{
        state: "visible",
        timeout: 10000
    }});

    await end.waitFor({{
        state: "visible",
        timeout: 10000
    }});


    // ========================================================
    // BEFORE VALUES
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("TIME VALUES BEFORE CHANGES");
    console.log("========================================");

    console.log(
        "START BEFORE = [" +
        await start.inputValue() +
        "]"
    );

    console.log(
        "END BEFORE = [" +
        await end.inputValue() +
        "]"
    );


    // ========================================================
    // JOB START
    // ========================================================

    if ({start_time!r}) {{

        console.log("");
        console.log("========================================");
        console.log("SETTING JOB START");
        console.log("========================================");

        console.log(
            "REQUESTED START = [{start_time}]"
        );

        await start.fill(
            {start_time!r}
        );

        console.log(
            "START IMMEDIATELY = [" +
            await start.inputValue() +
            "]"
        );

        await page.waitForTimeout(1500);

        console.log(
            "START AFTER 1.5 SEC = [" +
            await start.inputValue() +
            "]"
        );

    }} else {{

        console.log(
            "No Start time supplied - leaving unchanged."
        );

    }}


    // ========================================================
    // JOB END
    // ========================================================

    if ({end_time!r}) {{

        console.log("");
        console.log("========================================");
        console.log("SETTING JOB END");
        console.log("========================================");

        console.log(
            "REQUESTED END = [{end_time}]"
        );

        await end.fill(
            {end_time!r}
        );

        console.log(
            "END IMMEDIATELY = [" +
            await end.inputValue() +
            "]"
        );

        await page.waitForTimeout(500);

        console.log(
            "END AFTER 0.5 SEC = [" +
            await end.inputValue() +
            "]"
        );

    }} else {{

        console.log(
            "No End time supplied - leaving unchanged."
        );

    }}


    // ========================================================
    // STATUS
    // ========================================================

    console.log("");
    console.log(
        "Setting status: {status}"
    );

    const statusCombobox = page.getByRole(
        "combobox",
        {{ name: "Status *" }}
    );

    await statusCombobox.click();

    await page.waitForTimeout(500);

    const statusOption = page.locator(
        ".Select__option",
        {{ hasText: {status!r} }}
    ).last();

    await statusOption.click();

    await page.waitForTimeout(500);


    // ========================================================
    // CHARGE TYPE
    // ========================================================

    console.log(
        "Setting charge type: {charge_type}"
    );

    const chargeCombobox = page.getByRole(
        "combobox",
        {{ name: "Charge Type *" }}
    );

    await chargeCombobox.click();

    await page.waitForTimeout(500);

    const chargeOption = page.locator(
        ".Select__option",
        {{ hasText: {charge_type!r} }}
    ).last();

    await chargeOption.click();

    await page.waitForTimeout(500);


    // ========================================================
    // FINAL VERIFICATION BEFORE SAVE
    // ========================================================

    const finalStart = await page.locator(
        'input[name="actionarrivaldate_time"]'
    ).inputValue();

    const finalEnd = await page.locator(
        'input[name="actioncompletiondate_time"]'
    ).inputValue();


    console.log("");
    console.log("========================================");
    console.log("FINAL VALUES BEFORE SAVE");
    console.log("========================================");

    console.log(
        "FINAL START = [" +
        finalStart +
        "]"
    );

    console.log(
        "FINAL END = [" +
        finalEnd +
        "]"
    );


    // ========================================================
    // SAFETY CHECK
    // ========================================================

    if (
        {start_time!r}
        &&
        finalStart !== {start_time!r}
    ) {{

        throw new Error(
            "JOB START CHANGED BEFORE SAVE. " +
            "Requested=[" +
            {start_time!r} +
            "] Actual=[" +
            finalStart +
            "]"
        );

    }}


    if (
        {end_time!r}
        &&
        finalEnd !== {end_time!r}
    ) {{

        throw new Error(
            "JOB END CHANGED BEFORE SAVE. " +
            "Requested=[" +
            {end_time!r} +
            "] Actual=[" +
            finalEnd +
            "]"
        );

    }}


    // ========================================================
    // SAVE
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("SAVING WORKLOG");
    console.log("========================================");

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

    console.log(
        "Worklog saved."
    );


    return {{
        start: finalStart,
        end: finalEnd,
        url: page.url()
    }};
}}
"""

    temp_path = None

    try:

        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".js",
            prefix="halo_worklog_",
            delete=False,
            encoding="utf-8",
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

        # -----------------------------------------------------
        # ATTACH TO EDGE
        # -----------------------------------------------------

        attach()


        ticket = ""


        # -----------------------------------------------------
        # TICKET SELECTION
        # -----------------------------------------------------

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
                        f"{type_info} - "
                        f"{t['title']}"
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


        # -----------------------------------------------------
        # WORKLOG TEXT
        # -----------------------------------------------------

        worklog_text = ""


        while not worklog_text:

            worklog_text = input(
                "Worklog text (Required): "
            ).strip()


            if not worklog_text:

                print(
                    "Worklog text cannot be empty."
                )


        # -----------------------------------------------------
        # STATUS
        # -----------------------------------------------------

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
                        "Try again or enter custom text."
                    )


            elif not status_input:

                status = default_status

                break


            else:

                status = status_input

                break


        # -----------------------------------------------------
        # JOB TIMES
        # -----------------------------------------------------

        start_time = input(
            "Start time "
            "[Leave unchanged, e.g. 09:00]: "
        ).strip()


        end_time = input(
            "End time "
            "[Leave unchanged, e.g. 10:00]: "
        ).strip()


        # -----------------------------------------------------
        # CHARGE TYPE
        # -----------------------------------------------------

        charge_options = [
            "Project Work- Managed Services",
            "Research (work-specific)",
            "Professional Development",
            "Internal Work"
        ]


        default_charge = (
            "Internal Work"
        )


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

                    charge_type = (
                        charge_options[
                            int(sel) - 1
                        ]
                    )

                    break

                else:

                    print(
                        "Invalid selection. "
                        "Try again or enter custom text."
                    )


            elif not charge_input:

                charge_type = default_charge

                break


            else:

                charge_type = charge_input

                break


        # -----------------------------------------------------
        # OPEN TICKET
        # -----------------------------------------------------

        goto_ticket(
            ticket
        )


        # -----------------------------------------------------
        # RUN AUTOMATION
        # -----------------------------------------------------

        run_halo_automation(
            worklog_text,
            status,
            start_time,
            end_time,
            charge_type
        )


        # -----------------------------------------------------
        # FINAL SNAPSHOT
        # -----------------------------------------------------

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


if __name__ == "__main__":
    main()