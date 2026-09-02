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
    Remove local .playwright or .playwright-cli folders.
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


def attach():
    """
    Attach the CLI session to Microsoft Edge and open
    the Assigned Tickets view.
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
        url
    )

    time.sleep(1)


def parse_snapshot_tickets(snapshot_output):
    """
    Parse the playwright-cli snapshot text output into
    structured ticket dictionaries.
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


def run_halo_automation(
    worklog_text,
    status,
    start_time,
    end_time,
    charge_type
):
    """
    Run the HaloPSA Worklog automation.

    The Job Start and Job End fields use the actual HTML
    name attributes discovered from the live Halo page:

        Job Start:
            input[name="actionarrivaldate_time"]

        Job End:
            input[name="actioncompletiondate_time"]

    We deliberately do NOT use snapshot ref numbers,
    global input indexes, or the generated IDs containing
    "undefined".
    """

    js_code = f"""
async page => {{

    console.log("Opening Worklog...");


    // ---------------------------------------------------------
    // OPEN WORKLOG
    // ---------------------------------------------------------

    await page.getByRole(
        "button",
        {{
            name: "Worklog"
        }}
    ).click();


    await page.waitForTimeout(1500);


    // ---------------------------------------------------------
    // WORKLOG TEXT
    // ---------------------------------------------------------

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


    // ---------------------------------------------------------
    // STATUS
    // ---------------------------------------------------------

    console.log(
        "Setting status: {status}"
    );


    const statusCombobox =
        page.getByRole(
            "combobox",
            {{
                name: "Status *"
            }}
        );


    await statusCombobox.click();


    await page.waitForTimeout(800);


    try {{

        await page.getByText(
            {status!r},
            {{
                exact: true
            }}
        ).click();

    }} catch (e) {{

        await page.evaluate(
            targetText => {{

                const items =
                    Array.from(
                        document.querySelectorAll(
                            '.dropdown-item, [role="option"], li, div, .Select__option'
                        )
                    );


                const match =
                    items.find(
                        el =>
                            el.offsetParent !== null
                            &&
                            el.textContent.trim()
                                === targetText
                    );


                if (match) {{

                    match.click();

                }}

            }},
            {status!r}
        );
    }}


    // ---------------------------------------------------------
    // JOB START / END TIMES
    // ---------------------------------------------------------

    console.log("");
    console.log(
        "========================================"
    );
    console.log(
        "JOB START / END TIMES"
    );
    console.log(
        "========================================"
    );


    // These are the actual HTML name attributes discovered
    // from the live Halo page.
    //
    // Job Start:
    // actionarrivaldate_time
    //
    // Job End:
    // actioncompletiondate_time


    const startTimeInput =
        page.locator(
            'input[name="actionarrivaldate_time"]'
        );


    const endTimeInput =
        page.locator(
            'input[name="actioncompletiondate_time"]'
        );


    await startTimeInput.waitFor({{
        state: "visible",
        timeout: 10000
    }});


    await endTimeInput.waitFor({{
        state: "visible",
        timeout: 10000
    }});


    console.log(
        "Job Start current value:",
        await startTimeInput.inputValue()
    );


    console.log(
        "Job End current value:",
        await endTimeInput.inputValue()
    );


    // ---------------------------------------------------------
    // JOB START
    // ---------------------------------------------------------

    if ({start_time!r}) {{

        console.log(
            "Setting Job Start to:",
            {start_time!r}
        );


        await startTimeInput.click();


        await startTimeInput.fill(
            {start_time!r}
        );


        // Blur the field so Halo receives the change.
        await startTimeInput.press(
            "Tab"
        );


        await page.waitForTimeout(750);


        const actualStart =
            await startTimeInput.inputValue();


        console.log(
            "Job Start after update:",
            actualStart
        );


        if (
            actualStart !== {start_time!r}
        ) {{

            throw new Error(
                "Job Start failed. " +
                "Expected '" +
                {start_time!r} +
                "' but got '" +
                actualStart +
                "'."
            );
        }}


        console.log(
            "Job Start successfully updated."
        );

    }} else {{

        console.log(
            "Job Start left unchanged."
        );
    }}


    // ---------------------------------------------------------
    // JOB END
    // ---------------------------------------------------------

    if ({end_time!r}) {{

        console.log(
            "Setting Job End to:",
            {end_time!r}
        );


        await endTimeInput.click();


        await endTimeInput.fill(
            {end_time!r}
        );


        // Blur the field so Halo receives the change.
        await endTimeInput.press(
            "Tab"
        );


        await page.waitForTimeout(750);


        const actualEnd =
            await endTimeInput.inputValue();


        console.log(
            "Job End after update:",
            actualEnd
        );


        if (
            actualEnd !== {end_time!r}
        ) {{

            throw new Error(
                "Job End failed. " +
                "Expected '" +
                {end_time!r} +
                "' but got '" +
                actualEnd +
                "'."
            );
        }}


        console.log(
            "Job End successfully updated."
        );

    }} else {{

        console.log(
            "Job End left unchanged."
        );
    }}


    console.log(
        "========================================"
    );


    // ---------------------------------------------------------
    // FINAL TIME VERIFICATION
    // ---------------------------------------------------------

    console.log(
        "Final Job Start:",
        await startTimeInput.inputValue()
    );


    console.log(
        "Final Job End:",
        await endTimeInput.inputValue()
    );


    // ---------------------------------------------------------
    // CHARGE TYPE
    // ---------------------------------------------------------

    console.log(
        "Setting charge type: {charge_type}"
    );


    let chargeSelected = false;


    for (
        let attempt = 1;
        attempt <= 3;
        attempt++
    ) {{

        try {{

            const chargeTypeCombobox =
                page.getByRole(
                    "combobox",
                    {{
                        name: "Charge Type *"
                    }}
                );


            await chargeTypeCombobox.click();


            await page.waitForTimeout(800);


            const clicked =
                await page.evaluate(
                    targetText => {{

                        const selectors = [
                            ".Select__option",
                            ".dropdown-item",
                            '[role="option"]',
                            "li",
                            ".select2-results__option",
                            ".ng-option",
                            "div",
                            "span"
                        ];


                        for (
                            const selector
                            of selectors
                        ) {{

                            const items =
                                Array.from(
                                    document.querySelectorAll(
                                        selector
                                    )
                                );


                            const match =
                                items.find(
                                    el =>
                                        el.offsetParent !== null
                                        &&
                                        el.textContent.trim()
                                            === targetText
                                );


                            if (match) {{

                                match.click();

                                return true;
                            }}
                        }}


                        return false;

                    }},
                    {charge_type!r}
                );


            if (clicked) {{

                chargeSelected = true;

                break;
            }}

        }} catch (err) {{

            console.log(
                "Charge attempt " +
                attempt +
                " failed. Retrying..."
            );
        }}


        await page.waitForTimeout(1000);
    }}


    if (!chargeSelected) {{

        try {{

            await page.keyboard.type(
                {charge_type!r}
            );


            await page.keyboard.press(
                "Enter"
            );


            chargeSelected = true;

        }} catch (e) {{

            throw new Error(
                "Failed to select Charge Type: " +
                {charge_type!r}
            );
        }}
    }}


    await page.waitForTimeout(1000);


    // ---------------------------------------------------------
    // SAVE
    // ---------------------------------------------------------

    console.log(
        "Worklog fields populated. Saving..."
    );


    await page.evaluate(() => {{

        document.body.click();


        const saveBtn =
            Array.from(
                document.querySelectorAll(
                    'input[type="submit"], button'
                )
            ).find(
                el =>
                    el.value === "Save"
                    ||
                    el.textContent.trim() === "Save"
            );


        if (!saveBtn) {{

            throw new Error(
                "Save button not found in DOM."
            );
        }}


        saveBtn.click();
    }});


    await page.waitForTimeout(2000);


    console.log(
        "Worklog saved."
    );
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


def take_snapshot():

    print(
        "\nTaking final snapshot..."
    )

    run_cli(
        "snapshot"
    )


def main():

    print()
    print(
        "HaloPSA Worklog Automation"
    )
    print(
        "=========================="
    )


    try:

        attach()


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
        # NAVIGATE TO TICKET
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