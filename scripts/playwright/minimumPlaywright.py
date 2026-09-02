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


# ============================================================
# CONFIGURATION
# ============================================================

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

    os.chmod(
        path,
        stat.S_IWRITE
    )

    func(path)


def cleanup_local_artifacts():

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
# EDGE ATTACH
# ============================================================

def attach():

    print()
    print("========================================")
    print("ATTACHING TO EDGE")
    print("========================================")

    run_cli(
        "attach",
        "--extension=msedge",
        check=True
    )

    time.sleep(1)


# ============================================================
# NAVIGATE TO TICKET
# ============================================================

def goto_ticket(ticket):

    url = (
        f"{BASE_URL}"
        f"?id={ticket}"
        f"&showalltickettypes=1"
    )

    print()
    print(
        f"Opening ticket {ticket}..."
    )

    run_cli(
        "goto",
        url,
        check=True
    )

    time.sleep(1)


# ============================================================
# MINIMAL WORKLOG TEST
# ============================================================

def run_time_test(start_time, end_time):

    """
    IMPORTANT:

    This test intentionally does ONLY:

        1. Open Worklog
        2. Find Job Start
        3. Click Job Start
        4. Fill Job Start
        5. Find Job End
        6. Click Job End
        7. Fill Job End
        8. Verify both values
        9. Save

    It does NOT touch:

        - Worklog editor
        - Status
        - Charge Type

    This lets us compare the Python-generated run-code
    against the exact live test that previously worked.
    """

    js_code = f"""
async page => {{

    console.log("");
    console.log("========================================");
    console.log("MINIMAL HALO TIME TEST");
    console.log("========================================");


    // ========================================================
    // STEP 1 - OPEN WORKLOG
    // ========================================================

    console.log("");
    console.log("STEP 1: Opening Worklog");

    await page.getByRole(
        "button",
        {{
            name: "Worklog"
        }}
    ).click();

    await page.waitForTimeout(1500);


    // ========================================================
    // STEP 2 - LOCATE INPUTS
    // ========================================================

    console.log("");
    console.log("STEP 2: Locating time inputs");


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


    console.log(
        "Start count:",
        await start.count()
    );

    console.log(
        "End count:",
        await end.count()
    );

    console.log(
        "Start visible:",
        await start.isVisible()
    );

    console.log(
        "End visible:",
        await end.isVisible()
    );

    console.log(
        "Start enabled:",
        await start.isEnabled()
    );

    console.log(
        "End enabled:",
        await end.isEnabled()
    );


    // ========================================================
    // INITIAL VALUES
    // ========================================================

    const initialStart =
        await start.inputValue();

    const initialEnd =
        await end.inputValue();


    console.log("");
    console.log(
        "INITIAL START:",
        initialStart
    );

    console.log(
        "INITIAL END:",
        initialEnd
    );


    // ========================================================
    // STEP 3 - START
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("STEP 3: SETTING JOB START");
    console.log("========================================");

    console.log(
        "Requested Start:",
        {start_time!r}
    );


    console.log(
        "Start before click:",
        await start.inputValue()
    );


    await start.scrollIntoViewIfNeeded();

    console.log(
        "Start scrolled into view."
    );


    await start.click();

    console.log(
        "Start CLICK completed."
    );


    const startFocused =
        await start.evaluate(
            el => document.activeElement === el
        );


    console.log(
        "Start focused:",
        startFocused
    );


    console.log(
        "Start before fill:",
        await start.inputValue()
    );


    await start.fill(
        {start_time!r}
    );


    console.log(
        "Start immediately after fill:",
        await start.inputValue()
    );


    await page.waitForTimeout(100);


    console.log(
        "Start after 100ms:",
        await start.inputValue()
    );


    await page.waitForTimeout(1500);


    console.log(
        "Start after 1.5 sec:",
        await start.inputValue()
    );


    // ========================================================
    // STEP 4 - END
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("STEP 4: SETTING JOB END");
    console.log("========================================");

    console.log(
        "Requested End:",
        {end_time!r}
    );


    console.log(
        "End before click:",
        await end.inputValue()
    );


    await end.scrollIntoViewIfNeeded();

    await end.click();

    console.log(
        "End CLICK completed."
    );


    await end.fill(
        {end_time!r}
    );


    console.log(
        "End immediately after fill:",
        await end.inputValue()
    );


    await page.waitForTimeout(100);


    console.log(
        "End after 100ms:",
        await end.inputValue()
    );


    await page.waitForTimeout(1500);


    console.log(
        "End after 1.5 sec:",
        await end.inputValue()
    );


    // ========================================================
    // STEP 5 - FINAL VALUES
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("STEP 5: FINAL VALUES");
    console.log("========================================");


    const finalStart =
        await start.inputValue();

    const finalEnd =
        await end.inputValue();


    console.log(
        "FINAL START:",
        finalStart
    );

    console.log(
        "FINAL END:",
        finalEnd
    );


    // ========================================================
    // HARD VALIDATION
    // ========================================================

    if (
        {start_time!r}
        && finalStart !== {start_time!r}
    ) {{

        throw new Error(
            "JOB START FAILED. "
            + "Requested="
            + {start_time!r}
            + " Actual="
            + finalStart
        );

    }}


    if (
        {end_time!r}
        && finalEnd !== {end_time!r}
    ) {{

        throw new Error(
            "JOB END FAILED. "
            + "Requested="
            + {end_time!r}
            + " Actual="
            + finalEnd
        );

    }}


    // ========================================================
    // STEP 6 - SAVE
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("STEP 6: SAVING");
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


    console.log(
        "Save clicked."
    );


    await page.waitForTimeout(2000);


    console.log("");
    console.log("========================================");
    console.log("SAVE COMPLETE");
    console.log("========================================");


    return {{
        initialStart,
        initialEnd,
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
            prefix="halo_time_test_",
            delete=False,
            encoding="utf-8"
        ) as temp_file:

            temp_file.write(
                js_code
            )

            temp_path = temp_file.name


        print()
        print(
            "Running Playwright time test..."
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
# MAIN
# ============================================================

def main():

    print()
    print("========================================")
    print("HALO WORKLOG TIME TEST")
    print("========================================")


    try:

        # ------------------------------------------------------
        # ATTACH TO EXISTING EDGE
        # ------------------------------------------------------

        attach()


        # ------------------------------------------------------
        # TICKET
        # ------------------------------------------------------

        ticket = input(
            "\nEnter Ticket Number: "
        ).strip()


        while (
            not ticket
            or not ticket.isdigit()
        ):

            print(
                "Please enter a valid numeric ticket number."
            )

            ticket = input(
                "Enter Ticket Number: "
            ).strip()


        # ------------------------------------------------------
        # TIMES
        # ------------------------------------------------------

        start_time = input(
            "\nStart time "
            "(e.g. 08:00): "
        ).strip()


        end_time = input(
            "End time "
            "(e.g. 09:00): "
        ).strip()


        # ------------------------------------------------------
        # VALIDATE TIME FORMAT
        # ------------------------------------------------------

        if not re.match(
            r"^\d{{2}}:\d{{2}}$",
            start_time
        ):

            raise ValueError(
                "Start time must be HH:MM, e.g. 08:00"
            )


        if not re.match(
            r"^\d{{2}}:\d{{2}}$",
            end_time
        ):

            raise ValueError(
                "End time must be HH:MM, e.g. 09:00"
            )


        # ------------------------------------------------------
        # OPEN TICKET
        # ------------------------------------------------------

        goto_ticket(
            ticket
        )


        # ------------------------------------------------------
        # RUN MINIMAL TEST
        # ------------------------------------------------------

        run_time_test(
            start_time,
            end_time
        )


        print()
        print("========================================")
        print("TEST COMPLETED SUCCESSFULLY")
        print("========================================")


    except Exception as exc:

        print()
        print("========================================")
        print("TEST FAILED")
        print("========================================")

        print(
            str(exc)
        )

        sys.exit(1)


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()