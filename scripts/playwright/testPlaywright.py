import os
import platform
import shutil
import subprocess
import sys
import tempfile
import textwrap
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

    raise FileNotFoundError("playwright-cli was not found in PATH.")


CLI = find_playwright_cli()


def run_cli(*args, check=True):
    """
    Run playwright-cli safely across platforms.
    """
    if platform.system() == "Windows":
        # Quote arguments on Windows so cmd.exe doesn't split on '&' or spaces
        quoted_args = []
        for arg in args:
            s = str(arg)
            if not (s.startswith('"') and s.endswith('"')):
                s = f'"{s}"'
            quoted_args.append(s)

        cmd_str = f'"{CLI}" "--s={SESSION}" ' + " ".join(quoted_args)

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
        command = [CLI, f"--s={SESSION}", *[str(arg) for arg in args]]

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
            f"Playwright CLI command failed with exit code {result.returncode}"
        )

    return result.stdout


def cleanup_local_artifacts():
    """
    Remove local .playwright or session folders created in the working directory.
    """
    folders_to_remove = [".playwright", ".playwright-cli"]
    
    for folder in folders_to_remove:
        if os.path.exists(folder) and os.path.isdir(folder):
            try:
                shutil.rmtree(folder)
                print(f"Cleaned up local folder: {folder}")
            except OSError:
                pass


def attach():
    """
    Attach the CLI session to the already-open Microsoft Edge tab
    through the Playwright browser extension.
    """
    print("Attaching to Microsoft Edge...")
    run_cli(
        "attach",
        "--extension=msedge",
    )
    time.sleep(1)


def goto_ticket(ticket):
    """
    Navigate to the requested Halo ticket.
    """
    url = f"{BASE_URL}?id={ticket}&showalltickettypes=1"

    print(f"\nOpening ticket {ticket}...")
    run_cli("goto", url)
    time.sleep(1)


def run_halo_automation():
    """
    Run the actual HaloPSA Worklog automation using Playwright code.
    """
    js_code = textwrap.dedent(
        f"""
        async page => {{

            console.log("Opening Worklog...");

            await page.getByRole("button", {{
                name: "Worklog"
            }}).click();

            // Give the modal and system clock default values time to render
            await page.waitForTimeout(1500);

            // ---------------------------------------------------------
            // WORKLOG TEXT
            // ---------------------------------------------------------

            console.log("Entering worklog...");

            const editor = page.locator(
                '[contenteditable="true"]'
            ).first();

            await editor.waitFor({{
                state: "visible",
                timeout: 10000
            }});

            await editor.fill(
                {WORKLOG_TEXT!r}
            );

            // ---------------------------------------------------------
            // STATUS
            // ---------------------------------------------------------

            console.log("Setting status: {STATUS}");

            const status = page.getByRole(
                "combobox",
                {{ name: "Status *" }}
            );

            await status.click();

            await page.getByText(
                {STATUS!r},
                {{ exact: true }}
            ).click();

            // ---------------------------------------------------------
            // JOB START / END TIMES
            // ---------------------------------------------------------

            console.log("Setting Job Start: {START_TIME}");
            console.log("Setting Job End: {END_TIME}");

            /*
             * Target ONLY time inputs (prefilled with system time containing ':') 
             * while explicitly ignoring date inputs (which contain '-' or '/').
             */
            const timeInputIndexes = await page.locator("input").evaluateAll(inputs => {{
                const result = [];
                inputs.forEach((input, index) => {{
                    const value = input.value || "";
                    const type = (input.getAttribute("type") || "text").toLowerCase();
                    const style = window.getComputedStyle(input);
                    
                    if (style.display !== "none" && style.visibility !== "hidden" && (type === "text" || type === "time")) {{
                        // Must look like a time (has ':') and NOT a date (no '-' or '/')
                        if (value.includes(":") && !value.includes("-") && !value.includes("/")) {{
                            result.push(index);
                        }}
                    }}
                }});
                return result;
            }});

            console.log("Time input indices found:", timeInputIndexes);

            if (timeInputIndexes.length < 2) {{
                throw new Error(
                    "Could not find the two Halo Worklog time inputs. Found " + timeInputIndexes.length
                );
            }}

            const allInputs = page.locator("input");

            await allInputs
                .nth(timeInputIndexes[0])
                .fill("{START_TIME}");

            await allInputs
                .nth(timeInputIndexes[1])
                .fill("{END_TIME}");

            // ---------------------------------------------------------
            // CHARGE TYPE
            // ---------------------------------------------------------

            console.log(
                "Setting charge type: {CHARGE_TYPE}"
            );

            const chargeType = page.getByRole(
                "combobox",
                {{ name: "Charge Type *" }}
            );

            await chargeType.click();

            await page.getByText(
                {CHARGE_TYPE!r},
                {{ exact: true }}
            ).click();

            // ---------------------------------------------------------
            // FINAL CHECK
            // ---------------------------------------------------------

            console.log(
                "Worklog fields populated. Saving..."
            );

            await page.waitForTimeout(500);

            await page.getByRole(
                "button",
                {{ name: "Save", exact: true }}
            ).click();

            await page.waitForTimeout(1500);

            console.log("Worklog saved.");

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
            encoding="utf-8",
        ) as temp_file:

            temp_file.write(js_code)
            temp_path = temp_file.name

        print("\nRunning Halo automation...")

        run_cli(
            "run-code",
            f"--filename={temp_path}",
        )

    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


def take_snapshot():
    """
    Take a final snapshot so we can verify the result.
    """
    print("\nTaking final snapshot...")
    run_cli("snapshot")


def main():
    print()
    print("HaloPSA Worklog Automation")
    print("==========================")

    ticket = input("Ticket number: ").strip()

    if not ticket.isdigit():
        print("Invalid ticket number.")
        return

    try:
        attach()
        goto_ticket(ticket)
        run_halo_automation()
        take_snapshot()

        print()
        print("================================")
        print("Worklog automation completed.")
        print("================================")

    except Exception as exc:
        print()
        print("ERROR:")
        print(exc)
        sys.exit(1)

    finally:
        cleanup_local_artifacts()


if __name__ == "__main__":
    main()