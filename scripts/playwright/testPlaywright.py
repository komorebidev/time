import subprocess
import sys

SESSION = "halo"


def pw(*args):
    result = subprocess.run(
        ["playwright-cli", f"-s={SESSION}", *args],
        text=True,
        capture_output=True,
    )

    if result.stdout:
        print(result.stdout)

    if result.returncode != 0:
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        raise RuntimeError(f"Playwright CLI failed: {' '.join(args)}")

    return result.stdout


def attach():
    print("Attaching to Microsoft Edge...")
    
    result = subprocess.run(
        [
            "playwright-cli",
            "attach",
            "--extension=msedge",
            f"-s={SESSION}",
        ],
        text=True,
        capture_output=True,
    )

    print(result.stdout)

    if result.returncode != 0:
        print(result.stderr, file=sys.stderr)
        raise RuntimeError("Could not attach to Edge.")


def main():
    attach()

    ticket = input("Ticket number: ").strip()

    if not ticket.isdigit():
        print("Invalid ticket number.")
        return

    url = (
        f"https://support.eiresystems.com/"
        f"ticket?id={ticket}&showalltickettypes=1"
    )

    print(f"Opening ticket {ticket}...")
    pw("goto", url)

    print("Taking snapshot...")
    pw("snapshot")


if __name__ == "__main__":
    main()