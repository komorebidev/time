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


def run_halo_automation(worklog_text, status, start_time, end_time, charge_type):
    """
    Run the actual HaloPSA Worklog automation using Playwright code.
    """
    # Conditionally build the JS lines for filling start/end times only if provided
    start_fill_code = f'await allInputs.nth(timeInputIndexes[0]).fill({start_time!r});' if start_time else '// Start time left unchanged'
    end_fill_code = f'await allInputs.nth(timeInputIndexes[1]).fill({end_time!r});' if end_time else '// End time left unchanged'

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
                {worklog_text!r}
            );

            // ---------------------------------------------------------
            // STATUS
            // ---------------------------------------------------------

            console.log("Setting status: {status}");

            const statusCombobox = page.getByRole(
                "combobox",
                {{ name: "Status *" }}
            );

            await statusCombobox.click();

            await page.getByText(
                {status!r},
                {{ exact: true }}
            ).click();

            // ---------------------------------------------------------
            // JOB START / END TIMES
            // ---------------------------------------------------------

            console.log("Job Start Time input: {start_time if start_time else '(Leave unchanged)'}");
            console.log("Job End Time input: {end_time if end_time else '(Leave unchanged)'}");

            const timeInputIndexes = await page.locator("input").evaluateAll(inputs => {{
                const result = [];
                inputs.forEach((input, index) => {{
                    const value = input.value || "";
                    const type = (input.getAttribute("type") || "text").toLowerCase();
                    const style = window.getComputedStyle(input);
                    
                    if (style.display !== "none" && style.visibility !== "hidden" && (type === "text" || type === "time")) {{
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

            {start_fill_code}
            {end_fill_code}

            // ---------------------------------------------------------
            // CHARGE TYPE
            // ---------------------------------------------------------

            console.log(
                "Setting charge type: {charge_type}"
            );

            const chargeTypeCombobox = page.getByRole(
                "combobox",
                {{ name: "Charge Type *" }}
            );

            await chargeTypeCombobox.click();

            await page.getByText(
                {charge_type!r},
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

    # Interactive Inputs
    ticket = input("Ticket number: ").strip()
    if not ticket.isdigit():
        print("Invalid ticket number.")
        return

    # Worklog text is now fully required (no default)
    worklog_text = ""
    while not worklog_text:
        worklog_text = input("Worklog text (Required): ").strip()
        if not worklog_text:
            print("Worklog text cannot be empty. Please enter text.")

    default_status = "Completed (On Hold)"
    status = input(f"Status [{default_status}]: ").strip() or default_status

    # Leave start and end times blank by default to keep current values
    start_time = input("Start time [Leave unchanged, e.g. 09:00]: ").strip()
    end_time = input("End time [Leave unchanged, e.g. 10:00]: ").strip()

    default_charge = "Internal work"
    charge_type = input(f"Charge type [{default_charge}]: ").strip() or default_charge

    try:
        attach()
        goto_ticket(ticket)
        run_halo_automation(worklog_text, status, start_time, end_time, charge_type)
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