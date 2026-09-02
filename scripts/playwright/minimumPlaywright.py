import os
import platform
import shutil
import subprocess
import tempfile
import time


SESSION = "halo"
TICKET = "26229"
START_TIME = "08:00"

BASE_URL = "https://support.eiresystems.com/ticket"


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


def run_cli(*args):
    if platform.system() == "Windows":
        command = [
            CLI,
            f"--s={SESSION}",
            *[str(arg) for arg in args]
        ]

        print("\n>", " ".join(command))

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            shell=False,
        )

    else:
        command = [
            CLI,
            f"--s={SESSION}",
            *[str(arg) for arg in args]
        ]

        print("\n>", " ".join(command))

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            shell=False,
        )

    print(result.stdout)

    if result.returncode != 0:
        raise RuntimeError(
            f"playwright-cli failed with exit code {result.returncode}"
        )

    return result.stdout


def main():

    ticket_url = (
        f"{BASE_URL}"
        f"?id={TICKET}"
        f"&showalltickettypes=1"
    )

    js = f"""
async page => {{

    console.log("========================================");
    console.log("PYTHON -> PLAYWRIGHT START TEST");
    console.log("========================================");

    console.log("Opening ticket:");
    console.log("{ticket_url}");

    await page.goto("{ticket_url}");

    await page.waitForTimeout(1500);


    // --------------------------------------------------------
    // OPEN WORKLOG
    // --------------------------------------------------------

    console.log("");
    console.log("Opening Worklog...");

    const worklogButton = page.getByRole(
        "button",
        {{ name: "Worklog" }}
    );

    await worklogButton.click();

    await page.waitForTimeout(1500);


    // --------------------------------------------------------
    // LOCATE START FIELD
    // --------------------------------------------------------

    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );


    console.log("");
    console.log("Start field count:", await start.count());
    console.log("Start field visible:", await start.isVisible());
    console.log("Start field enabled:", await start.isEnabled());
    console.log("Start field ID:", await start.getAttribute("id"));
    console.log("Start field name:", await start.getAttribute("name"));
    console.log("Start field type:", await start.getAttribute("type"));


    const before = await start.inputValue();

    console.log("");
    console.log("BEFORE:", before);


    // --------------------------------------------------------
    // CLICK
    // --------------------------------------------------------

    console.log("");
    console.log("CLICKING JOB START...");

    await start.click();

    console.log(
        "Focused:",
        await start.evaluate(
            el => document.activeElement === el
        )
    );

    console.log(
        "Active element:",
        await page.evaluate(
            () => ({{
                id: document.activeElement?.id,
                name: document.activeElement?.getAttribute("name"),
                type: document.activeElement?.type,
                value: document.activeElement?.value
            }})
        )
    );


    // --------------------------------------------------------
    // FILL
    // --------------------------------------------------------

    console.log("");
    console.log("FILLING JOB START WITH:");
    console.log("{START_TIME}");

    await start.fill("{START_TIME}");


    const afterFill = await start.inputValue();

    console.log("");
    console.log("AFTER FILL:", afterFill);


    // --------------------------------------------------------
    // WAIT
    // --------------------------------------------------------

    await page.waitForTimeout(1500);

    const afterWait = await start.inputValue();

    console.log(
        "AFTER 1.5 SEC:",
        afterWait
    );


    // --------------------------------------------------------
    // RETURN DIAGNOSTIC
    // --------------------------------------------------------

    return {{
        before,
        afterFill,
        afterWait,

        count: await start.count(),
        visible: await start.isVisible(),
        enabled: await start.isEnabled(),

        id: await start.getAttribute("id"),
        name: await start.getAttribute("name"),
        type: await start.getAttribute("type"),

        focused: await start.evaluate(
            el => document.activeElement === el
        )
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

            f.write(js)
            temp_path = f.name


        run_cli(
            "run-code",
            f"--filename={temp_path}"
        )


    finally:

        if temp_path and os.path.exists(temp_path):

            try:
                os.remove(temp_path)
            except OSError:
                pass


if __name__ == "__main__":
    main()