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
    "?area=1&mainview=team&viewid=2&selid=75"
    "&sellevel=2&selparentid=engineers%20tky"
)

SESSION = "halo"


# ============================================================
# PLAYWRIGHT CLI
# ============================================================

def find_playwright_cli():
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
    Run playwright-cli against the existing Edge session.
    """

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

    if check and result.returncode != 0:
        raise RuntimeError(
            "Playwright CLI command failed "
            f"with exit code {result.returncode}"
        )

    return result.stdout


# ============================================================
# CLEANUP
# ============================================================

def remove_readonly(func, path, excinfo):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        pass


def cleanup_local_artifacts():
    folders_to_remove = [
        ".playwright",
        ".playwright-cli"
    ]

    for folder in folders_to_remove:

        if os.path.isdir(folder):

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
                            f"{folder}: {e}"
                        )
                    else:
                        time.sleep(0.5)


atexit.register(cleanup_local_artifacts)


# ============================================================
# EDGE
# ============================================================

def attach():
    """
    Attach to the existing Microsoft Edge session.
    """

    print("\nAttaching to Microsoft Edge...")

    run_cli(
        "attach",
        "--extension=msedge",
        check=True
    )

    time.sleep(1)

    print("\nOpening Assigned Tickets view...")

    run_cli(
        "goto",
        ASSIGNED_TICKETS_URL,
        check=True
    )

    time.sleep(2)


def goto_ticket(ticket):
    """
    Navigate to the requested Halo ticket.
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
                        and not val.startswith("[")
                    ):
                        clean_lines.append(val)

    for i, l in enumerate(clean_lines):

        if re.match(
            r'^00\d{5}$',
            l
        ):

            ticket_id = l

        elif (
            "/" in l
            and "EIRE" in l
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
                ticket_name = clean_lines[i - 1]

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
    Complete HaloPSA worklog automation.

    IMPORTANT:

    Job Start is intentionally handled FIRST.

    Exact sequence:

        1. Open Worklog
        2. Find Job Start
        3. Click Job Start
        4. Fill Job Start

    Only AFTER that do we touch the other fields.
    """

    # --------------------------------------------------------
    # Safely encode Python values into JavaScript literals
    # --------------------------------------------------------

    worklog_js = repr(worklog_text)
    status_js = repr(status)
    start_js = repr(start_time)
    end_js = repr(end_time)
    charge_js = repr(charge_type)

    js_code = f"""
async page => {{

    console.log("");
    console.log("========================================");
    console.log("HALO WORKLOG AUTOMATION");
    console.log("========================================");


    // ========================================================
    // 1. OPEN WORKLOG
    // ========================================================

    console.log("");
    console.log("STEP 1: OPEN WORKLOG");

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


    // ========================================================
    // 2. FIND JOB START
    // ========================================================

    console.log("");
    console.log("STEP 2: FIND JOB START");

    const startTimeInput = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    await startTimeInput.waitFor({{
        state: "visible",
        timeout: 10000
    }});

    const startCount =
        await startTimeInput.count();

    const startBefore =
        await startTimeInput.inputValue();

    console.log(
        "Job Start count:",
        startCount
    );

    console.log(
        "Job Start before:",
        startBefore
    );

    console.log(
        "Job Start ID:",
        await startTimeInput.getAttribute("id")
    );

    console.log(
        "Job Start name:",
        await startTimeInput.getAttribute("name")
    );

    console.log(
        "Job Start type:",
        await startTimeInput.getAttribute("type")
    );


    // ========================================================
    // 3. CLICK JOB START
    // ========================================================

    console.log("");
    console.log("STEP 3: CLICK JOB START");

    await startTimeInput.click();

    const focused =
        await startTimeInput.evaluate(
            el => document.activeElement === el
        );

    console.log(
        "Job Start focused:",
        focused
    );

    if (!focused) {{
        throw new Error(
            "Job Start field was clicked but did not receive focus."
        );
    }}


    // ========================================================
    // 4. FILL JOB START
    // ========================================================

    if ({start_js}) {{

        console.log("");
        console.log("STEP 4: FILL JOB START");

        console.log(
            "Requested Job Start:",
            {start_js}
        );

        await startTimeInput.fill(
            {start_js}
        );

        const startAfterFill =
            await startTimeInput.inputValue();

        console.log(
            "Job Start immediately after fill:",
            startAfterFill
        );

        if (startAfterFill !== {start_js}) {{
            throw new Error(
                "Job Start did not accept the requested value. " +
                "Expected: " + {start_js} +
                ", Actual: " + startAfterFill
            );
        }}

        // Give Halo time to process the value,
        // but DO NOT modify the field again.
        await page.waitForTimeout(500);

        const startAfterWait =
            await startTimeInput.inputValue();

        console.log(
            "Job Start after 500ms:",
            startAfterWait
        );

        if (startAfterWait !== {start_js}) {{
            throw new Error(
                "Job Start changed unexpectedly after fill. " +
                "Expected: " + {start_js} +
                ", Actual: " + startAfterWait
            );
        }}

    }} else {{

        console.log(
            "No Job Start supplied. Leaving unchanged."
        );

    }}


    // ========================================================
    // ONLY NOW TOUCH OTHER FIELDS
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("JOB START SUCCESSFULLY PROCESSED");
    console.log("NOW PROCESSING OTHER FIELDS");
    console.log("========================================");


    // ========================================================
    // WORKLOG TEXT
    // ========================================================

    console.log("");
    console.log("Entering worklog...");

    const editor = page.locator(
        '[contenteditable="true"]'
    ).first();

    await editor.waitFor({{
        state: "visible",
        timeout: 10000
    }});

    await editor.fill(
        {worklog_js}
    );


    // ========================================================
    // STATUS
    // ========================================================

    console.log("");
    console.log(
        "Setting status:",
        {status_js}
    );

    const statusCombobox = page.getByRole(
        "combobox",
        {{
            name: "Status *"
        }}
    );

    await statusCombobox.click();

    await page.waitForTimeout(500);


    const statusOption = page.locator(
        ".Select__option",
        {{
            hasText: {status_js}
        }}
    ).last();


    await statusOption.waitFor({{
        state: "visible",
        timeout: 10000
    }});

    await statusOption.click();

    await page.waitForTimeout(500);


    // ========================================================
    // JOB END
    // ========================================================

    if ({end_js}) {{

        console.log("");
        console.log("Setting Job End:");

        console.log(
            {end_js}
        );

        const endTimeInput = page.locator(
            'input[name="actioncompletiondate_time"]'
        );

        await endTimeInput.waitFor({{
            state: "visible",
            timeout: 10000
        }});

        await endTimeInput.click();

        await endTimeInput.fill(
            {end_js}
        );

        const endAfterFill =
            await endTimeInput.inputValue();

        console.log(
            "Job End after fill:",
            endAfterFill
        );

        if (endAfterFill !== {end_js}) {{
            throw new Error(
                "Job End did not accept the requested value. " +
                "Expected: " + {end_js} +
                ", Actual: " + endAfterFill
            );
        }}

        await endTimeInput.press("Tab");

        await page.waitForTimeout(500);

    }} else {{

        console.log(
            "No Job End supplied. Leaving unchanged."
        );

    }}


    // ========================================================
    // CHARGE TYPE
    // ========================================================

    console.log("");
    console.log(
        "Setting charge type:",
        {charge_js}
    );

    const chargeCombobox = page.getByRole(
        "combobox",
        {{
            name: "Charge Type *"
        }}
    );

    await chargeCombobox.click();

    await page.waitForTimeout(500);


    const chargeOption = page.locator(
        ".Select__option",
        {{
            hasText: {charge_js}
        }}
    ).last();


    await chargeOption.waitFor({{
        state: "visible",
        timeout: 10000
    }});

    await chargeOption.click();

    await page.waitForTimeout(500);


    // ========================================================
    // FINAL VERIFICATION
    // ========================================================

    const finalStart =
        await startTimeInput.inputValue();

    const endTimeInputFinal = page.locator(
        'input[name="actioncompletiondate_time"]'
    );

    const finalEnd =
        await endTimeInputFinal.inputValue();


    console.log("");
    console.log("========================================");
    console.log("FINAL VALUES BEFORE SAVE");
    console.log("========================================");

    console.log(
        "Job Start:",
        finalStart
    );

    console.log(
        "Job End:",
        finalEnd
    );


    // If a Start time was requested, make absolutely
    // sure it survived all the other field operations.

    if ({start_js}) {{

        if (finalStart !== {start_js}) {{
            throw new Error(
                "Job Start was correct initially but changed " +
                "before Save. Expected: " +
                {start_js} +
                ", Actual: " +
                finalStart
            );
        }}
    }}


    // ========================================================
    // SAVE
    // ========================================================

    console.log("");
    console.log("Saving worklog...");

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
        startBefore,
        finalStart,
        finalEnd,
        urlAfterSave: page.url()
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
            encoding="utf-8"
        ) as temp_file:

            temp_file.write(js_code)
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
                os.remove(temp_path)
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

        # ------------------------------------------------------
        # ATTACH TO EDGE
        # ------------------------------------------------------

        attach()


        # ------------------------------------------------------
        # SELECT TICKET
        # ------------------------------------------------------

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


        # ------------------------------------------------------
        # VALIDATE TICKET
        # ------------------------------------------------------

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


        # ------------------------------------------------------
        # WORKLOG TEXT
        # ------------------------------------------------------

        worklog_text = ""

        while not worklog_text:

            worklog_text = input(
                "Worklog text (Required): "
            ).strip()

            if not worklog_text:

                print(
                    "Worklog text cannot be empty."
                )


        # ------------------------------------------------------
        # STATUS
        # ------------------------------------------------------

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


        # ------------------------------------------------------
        # TIMES
        # ------------------------------------------------------

        start_time = input(
            "Start time "
            "(e.g. 08:00): "
        ).strip()

        end_time = input(
            "End time "
            "(e.g. 09:00): "
        ).strip()


        # ------------------------------------------------------
        # CHARGE TYPE
        # ------------------------------------------------------

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
                        "Try again or enter custom text."
                    )


            elif not charge_input:

                charge_type = default_charge
                break


            else:

                charge_type = charge_input
                break


        # ------------------------------------------------------
        # OPEN TICKET
        # ------------------------------------------------------

        goto_ticket(
            ticket
        )


        # ------------------------------------------------------
        # RUN AUTOMATION
        # ------------------------------------------------------

        run_halo_automation(
            worklog_text,
            status,
            start_time,
            end_time,
            charge_type
        )


        # ------------------------------------------------------
        # FINAL SNAPSHOT
        # ------------------------------------------------------

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