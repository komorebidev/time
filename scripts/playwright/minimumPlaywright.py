import os
import sys
import platform
import shutil
import subprocess
import tempfile
import time
import re


BASE_URL = "https://support.eiresystems.com/ticket"
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

        cmd = (
            f'"{CLI}" "--s={SESSION}" '
            + " ".join(quoted_args)
        )

        print()
        print("> " + cmd)

        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            shell=True
        )

    else:

        command = [
            CLI,
            f"--s={SESSION}",
            *[str(x) for x in args]
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
            shell=False
        )

    if result.stdout:
        print(result.stdout)

    if check and result.returncode != 0:
        raise RuntimeError(
            f"Playwright CLI failed with exit code "
            f"{result.returncode}"
        )

    return result.stdout


# ============================================================
# NAVIGATION
# ============================================================

def goto_ticket(ticket):

    url = (
        f"{BASE_URL}"
        f"?id={ticket}"
        f"&showalltickettypes=1"
    )

    print()
    print(f"Opening ticket {ticket}...")

    run_cli(
        "goto",
        url,
        check=True
    )

    time.sleep(1)


# ============================================================
# TIME VALIDATION
# ============================================================

def validate_time(value):

    value = value.strip()

    # Correct HH:MM validation.
    # Allows 00:00 through 23:59.

    if not re.fullmatch(
        r"(?:[01]\d|2[0-3]):[0-5]\d",
        value
    ):
        raise ValueError(
            f"Invalid time '{value}'. "
            "Use HH:MM, e.g. 08:00"
        )

    return value


# ============================================================
# TEST
# ============================================================

def run_test(ticket, requested_start, requested_end):

    requested_start = validate_time(requested_start)
    requested_end = validate_time(requested_end)

    js_code = f"""
async page => {{

    console.log("");
    console.log("========================================");
    console.log("HALO JOB START DIAGNOSTIC");
    console.log("========================================");


    // ========================================================
    // OPEN WORKLOG
    // ========================================================

    console.log("");
    console.log("1. Opening Worklog...");

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
    // LOCATE ACTUAL TIME INPUTS
    // ========================================================

    console.log("");
    console.log("2. Locating actual Halo time inputs...");


    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );


    console.log(
        "Start count:",
        await start.count()
    );

    console.log(
        "End count:",
        await end.count()
    );


    if (await start.count() !== 1) {{
        throw new Error(
            "Expected exactly one Job Start input."
        );
    }}

    if (await end.count() !== 1) {{
        throw new Error(
            "Expected exactly one Job End input."
        );
    }}


    // ========================================================
    // INSPECT START FIELD
    // ========================================================

    console.log("");
    console.log("3. Inspecting Job Start field...");


    const startInfo = await start.evaluate(el => ({{
        id: el.id,
        name: el.getAttribute("name"),
        type: el.getAttribute("type"),
        value: el.value,
        disabled: el.disabled,
        readOnly: el.readOnly,
        visible:
            !!(el.offsetWidth || el.offsetHeight),
        outerHTML: el.outerHTML
    }}));


    console.log(
        "START INFO:",
        JSON.stringify(startInfo)
    );


    // ========================================================
    // INITIAL VALUES
    // ========================================================

    const initialStart =
        await start.inputValue();

    const initialEnd =
        await end.inputValue();


    console.log("");
    console.log("4. Initial values:");

    console.log(
        "Job Start:",
        initialStart
    );

    console.log(
        "Job End:",
        initialEnd
    );


    // ========================================================
    // CLICK START
    // ========================================================

    console.log("");
    console.log("5. CLICKING JOB START...");


    await start.scrollIntoViewIfNeeded();

    await start.click();


    const focusInfo = await start.evaluate(el => ({{
        active:
            document.activeElement === el,

        activeId:
            document.activeElement?.id || null,

        activeName:
            document.activeElement?.getAttribute("name") || null,

        activeType:
            document.activeElement?.getAttribute("type") || null
    }}));


    console.log(
        "FOCUS AFTER CLICK:",
        JSON.stringify(focusInfo)
    );


    if (!focusInfo.active) {{
        console.log(
            "WARNING: Job Start did not become "
            "document.activeElement."
        );
    }} else {{
        console.log(
            "SUCCESS: Job Start is focused."
        );
    }}


    // ========================================================
    // FILL START
    // ========================================================

    console.log("");
    console.log(
        "6. FILLING JOB START:",
        {requested_start!r}
    );


    await start.fill(
        {requested_start!r}
    );


    const afterFill =
        await start.inputValue();


    console.log(
        "VALUE IMMEDIATELY AFTER FILL:",
        afterFill
    );


    // ========================================================
    // VERIFY DOM VALUE
    // ========================================================

    if (afterFill !== {requested_start!r}) {{
        throw new Error(
            "Job Start did not contain the requested "
            "value immediately after fill. " +
            "Expected " + {requested_start!r} +
            " but got " + afterFill
        );
    }}


    // ========================================================
    // WAIT
    // ========================================================

    console.log("");
    console.log(
        "7. Waiting 1.5 seconds..."
    );

    await page.waitForTimeout(1500);


    const afterWait =
        await start.inputValue();


    console.log(
        "VALUE AFTER 1.5 SEC:",
        afterWait
    );


    // ========================================================
    // SET END
    // ========================================================

    console.log("");
    console.log(
        "8. Setting Job End:",
        {requested_end!r}
    );


    await end.click();

    await end.fill(
        {requested_end!r}
    );


    const afterEnd =
        await end.inputValue();


    console.log(
        "Job End after fill:",
        afterEnd
    );


    // ========================================================
    // CHECK START AGAIN
    // ========================================================

    const startAfterEnd =
        await start.inputValue();


    console.log(
        "Job Start after setting End:",
        startAfterEnd
    );


    // ========================================================
    // TAB AWAY FROM END
    // ========================================================

    await end.press("Tab");

    await page.waitForTimeout(500);


    const afterTabStart =
        await start.inputValue();

    const afterTabEnd =
        await end.inputValue();


    console.log("");
    console.log("9. AFTER TAB:");

    console.log(
        "Job Start:",
        afterTabStart
    );

    console.log(
        "Job End:",
        afterTabEnd
    );


    // ========================================================
    // FINAL PRE-SAVE CHECK
    // ========================================================

    console.log("");
    console.log("========================================");
    console.log("10. FINAL PRE-SAVE CHECK");
    console.log("========================================");


    const finalStart =
        await start.inputValue();

    const finalEnd =
        await end.inputValue();


    console.log(
        "Expected Start:",
        {requested_start!r}
    );

    console.log(
        "Actual Start:",
        finalStart
    );

    console.log(
        "Expected End:",
        {requested_end!r}
    );

    console.log(
        "Actual End:",
        finalEnd
    );


    if (finalStart !== {requested_start!r}) {{
        throw new Error(
            "JOB START FAILED BEFORE SAVE. " +
            "Expected " + {requested_start!r} +
            " but got " + finalStart
        );
    }}


    if (finalEnd !== {requested_end!r}) {{
        throw new Error(
            "JOB END FAILED BEFORE SAVE. " +
            "Expected " + {requested_end!r} +
            " but got " + finalEnd
        );
    }}


    // ========================================================
    // SAVE
    // ========================================================

    console.log("");
    console.log("11. Saving...");


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
    console.log("SAVE COMPLETE");
    console.log("========================================");

    console.log(
        "URL:",
        page.url()
    );


    return {{
        initialStart,
        initialEnd,
        afterFill,
        afterWait,
        afterEnd,
        afterTabStart,
        afterTabEnd,
        finalStart,
        finalEnd,
        url: page.url()
    }};
}}
"""


    temp_path = None

    try:

        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".js",
            prefix="halo_start_test_",
            delete=False,
            encoding="utf-8"
        ) as f:

            f.write(js_code)
            temp_path = f.name


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
# MAIN
# ============================================================

def main():

    print()
    print("========================================")
    print("HALO JOB START TEST")
    print("========================================")


    try:

        ticket = input(
            "\nEnter Ticket Number: "
        ).strip()


        if not ticket.isdigit():

            raise ValueError(
                "Ticket number must be numeric."
            )


        start = input(
            "Start time (e.g. 08:00): "
        ).strip()


        end = input(
            "End time (e.g. 09:00): "
        ).strip()


        # IMPORTANT:
        # validate_time() is called here using the
        # exact HH:MM regex above.

        start = validate_time(start)
        end = validate_time(end)


        print()
        print("========================================")
        print("TEST PARAMETERS")
        print("========================================")

        print(
            "Ticket:",
            ticket
        )

        print(
            "Requested Start:",
            start
        )

        print(
            "Requested End:",
            end
        )


        goto_ticket(ticket)

        run_test(
            ticket,
            start,
            end
        )


        print()
        print("========================================")
        print("TEST PASSED")
        print("========================================")


    except Exception as exc:

        print()
        print("========================================")
        print("TEST FAILED")
        print("========================================")

        print(exc)

        sys.exit(1)


if __name__ == "__main__":
    main()