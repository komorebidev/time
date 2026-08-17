import os
import sys
import atexit
import json
import platform
import re
import shutil
import stat
import subprocess
import tempfile
import textwrap
import time

BASE_URL = "https://support.eiresystems.com/ticket"
SESSION = "halo"


def find_playwright_cli():
    """
    Find playwright-cli without hard-coding the Windows username.
    Works on Windows and macOS.
    """
    cli = shutil.which("playwright-cli")
    if cli:
        return cli

    cli = shutil.which("playwright-cli.cmd")
    if cli:
        return cli

    raise FileNotFoundError(
        "playwright-cli was not found in PATH. "
        "Make sure @playwright/cli is installed."
    )


CLI = find_playwright_cli()


def run_cli(*args, check=True):
    """
    Run playwright-cli safely across Windows and macOS/Linux.
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
            f"Playwright CLI command failed with exit code {result.returncode}"
        )

    return result.stdout


def remove_readonly(func, path, excinfo):
    """
    Error handler for shutil.rmtree.
    Mainly useful on Windows.
    """
    try:
        os.chmod(path, stat.S_IWRITE)
    except OSError:
        pass

    func(path)


def cleanup_local_artifacts():
    """
    Remove local Playwright folders created in the working directory.

    NOTE:
    This does not remove the Playwright CLI's global installation.
    """
    folders_to_remove = [
        ".playwright",
        ".playwright-cli"
    ]

    for folder in folders_to_remove:

        if os.path.exists(folder) and os.path.isdir(folder):

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


atexit.register(cleanup_local_artifacts)


def attach():
    """
    Attach the CLI session to Microsoft Edge.
    """
    print("\nAttaching to Microsoft Edge...")

    run_cli(
        "attach",
        "--extension=msedge",
        check=True
    )

    time.sleep(1)


def goto_ticket(ticket):
    """
    Navigate directly to the requested Halo ticket.
    """
    url = (
        f"{BASE_URL}"
        f"?id={ticket}"
        f"&showalltickettypes=1"
    )

    print(f"\nOpening ticket {ticket}...")

    run_cli(
        "goto",
        url
    )

    time.sleep(1)


def scrape_ticket_options():
    """
    Intelligently parse ticket IDs, titles, dates, and types
    from the playwright-cli snapshot output.
    """

    print(
        "\nTaking snapshot to extract tickets and metadata..."
    )

    output = run_cli(
        "snapshot",
        check=True
    )

    tickets = []
    seen_ids = set()
    lines = output.splitlines()

    def process_block(ticket_id, block_lines):

        if not ticket_id or ticket_id in seen_ids:
            return

        seen_ids.add(ticket_id)

        title = f"Ticket {ticket_id}"
        date_str = ""
        ticket_type = ""

        # ---------------------------------------------------------
        # FIRST PASS
        # Determine ticket type and date
        # ---------------------------------------------------------

        for line in block_lines:

            if re.search(
                r'\d{1,2}/\d{1,2}/\d{4}\s+\d{2}:\d{2}',
                line
            ):

                date_match = re.search(
                    r'([\d/]+\s+[\d:]+)',
                    line
                )

                if date_match:
                    date_str = date_match.group(1)

            if any(
                t in line
                for t in [
                    "Service Request",
                    "Incident",
                    "Project Support"
                ]
            ):

                for t in [
                    "Service Request",
                    "Incident",
                    "Project Support"
                ]:

                    if t in line:
                        ticket_type = t
                        break

        # ---------------------------------------------------------
        # COLLECT CANDIDATE TEXT
        # ---------------------------------------------------------

        ordered_candidates = []

        for line in block_lines:

            # Skip duration lines
            if re.search(
                r'^\s*-\s*(?:generic|text):\s*\d+:\d+\s*$',
                line
            ):
                continue

            # Skip dates
            if re.search(
                r'\d{1,2}/\d{1,2}/\d{4}\s+\d{2}:\d{2}',
                line
            ):
                continue

            # Skip ticket types
            if any(
                t in line
                for t in [
                    "Service Request",
                    "Incident",
                    "Project Support"
                ]
            ):
                continue

            clean_line = re.sub(
                r'\[ref=e\d+\]',
                '',
                line
            )

            clean_line = re.sub(
                r'^\s*-\s*(?:generic|text)?\s*(?:"[^"]*")?\s*:\s*',
                '',
                clean_line
            )

            clean_line = re.sub(
                r'^\s*-\s*(?:generic|text)\b',
                '',
                clean_line
            )

            clean_line = clean_line.strip('" :')

            if not clean_line or len(clean_line) < 2:
                continue

            lower_line = clean_line.lower()

            if re.match(
                r'^00\d{5}$',
                clean_line
            ):
                continue

            if re.match(
                r'^\d+:\d+$',
                clean_line
            ):
                continue

            if any(
                kw in lower_line
                for kw in [
                    'cursor=pointer',
                    'checkbox',
                    'bulk select',
                    'available',
                    'on hold',
                    'completed',
                    'low',
                    'medium',
                    'high'
                ]
            ):
                continue

            if len(clean_line) <= 3 and clean_line.isupper():
                continue

            ordered_candidates.append(
                clean_line
            )

        # ---------------------------------------------------------
        # POSITIONAL / TYPE-AWARE RESOLUTION
        # ---------------------------------------------------------

        if ticket_type == "Project Support":

            path_candidates = [
                c for c in ordered_candidates
                if '/' in c
            ]

            if path_candidates:
                title = path_candidates[0]

            elif ordered_candidates:
                title = ordered_candidates[-1]

        else:

            non_path_candidates = [
                c for c in ordered_candidates
                if '/' not in c
            ]

            if non_path_candidates:
                title = non_path_candidates[-1]

            elif ordered_candidates:
                title = ordered_candidates[-1]

        tickets.append({
            "id": ticket_id,
            "title": title,
            "date": date_str,
            "type": ticket_type
        })

    active_id = None
    block_buffer = []

    for line in lines:

        id_match = re.search(
            r'"(00\d{5})"',
            line
        )

        if id_match:

            if active_id:
                process_block(
                    active_id,
                    block_buffer
                )

            active_id = id_match.group(1)
            block_buffer = []

        elif active_id:

            block_buffer.append(line)

            if (
                "cursor=pointer" in line
                and len(block_buffer) > 4
            ):

                process_block(
                    active_id,
                    block_buffer
                )

                active_id = None
                block_buffer = []

    if active_id:
        process_block(
            active_id,
            block_buffer
        )

    return tickets


def run_halo_automation(
    worklog_text,
    status,
    start_time,
    end_time,
    charge_type
):
    """
    Run the HaloPSA Worklog automation using Playwright.
    """

    start_fill_code = (
        f"await allInputs.nth(timeInputIndexes[0]).fill({start_time!r});"
        if start_time
        else "// Start time left unchanged"
    )

    end_fill_code = (
        f"await allInputs.nth(timeInputIndexes[1]).fill({end_time!r});"
        if end_time
        else "// End time left unchanged"
    )

    js_code = textwrap.dedent(
        f"""
        async page => {{

            console.log("Opening Worklog...");

            await page.getByRole("button", {{
                name: "Worklog"
            }}).click();

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

            console.log(
                "Setting status: {status}"
            );

            const statusCombobox = page.getByRole(
                "combobox",
                {{ name: "Status *" }}
            );

            await statusCombobox.click();

            // FIX:
            // Use the dropdown option instead of getByText().
            // This prevents strict-mode violations when the same
            // status text exists elsewhere on the page.

            await page.getByRole(
                "option",
                {{
                    name: {status!r},
                    exact: true
                }}
            ).click();


            // ---------------------------------------------------------
            // JOB START / END TIMES
            // ---------------------------------------------------------

            console.log(
                "Job Start Time input: "
                + "{start_time if start_time else '(Leave unchanged)'}"
            );

            console.log(
                "Job End Time input: "
                + "{end_time if end_time else '(Leave unchanged)'}"
            );

            const timeInputIndexes =
                await page.locator("input").evaluateAll(
                    inputs => {{

                        const result = [];

                        inputs.forEach(
                            (input, index) => {{

                                const value =
                                    input.value || "";

                                const type =
                                    (
                                        input.getAttribute(
                                            "type"
                                        ) || "text"
                                    ).toLowerCase();

                                const style =
                                    window.getComputedStyle(
                                        input
                                    );

                                if (
                                    style.display !== "none"
                                    &&
                                    style.visibility !== "hidden"
                                    &&
                                    (
                                        type === "text"
                                        ||
                                        type === "time"
                                    )
                                ) {{

                                    if (
                                        value.includes(":")
                                        &&
                                        !value.includes("-")
                                        &&
                                        !value.includes("/")
                                    ) {{

                                        result.push(index);

                                    }}
                                }}
                            }}
                        );

                        return result;
                    }}
                );

            console.log(
                "Time input indices found:",
                timeInputIndexes
            );

            if (timeInputIndexes.length < 2) {{

                throw new Error(
                    "Could not find the two Halo Worklog "
                    + "time inputs. Found "
                    + timeInputIndexes.length
                );

            }}

            const allInputs =
                page.locator("input");


            {start_fill_code}

            {end_fill_code}


            // ---------------------------------------------------------
            // CHARGE TYPE
            // ---------------------------------------------------------

            console.log(
                "Setting charge type: {charge_type}"
            );

            const chargeTypeCombobox =
                page.getByRole(
                    "combobox",
                    {{ name: "Charge Type *" }}
                );

            await chargeTypeCombobox.click();

            // FIX:
            // Use the dropdown option instead of getByText().

            await page.getByRole(
                "option",
                {{
                    name: {charge_type!r},
                    exact: true
                }}
            ).click();


            // ---------------------------------------------------------
            // FINAL CHECK
            // ---------------------------------------------------------

            console.log(
                "Worklog