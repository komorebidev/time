import os
import sys
import re
import base64
from datetime import datetime
from azure.communication.email import EmailClient
from icalendar import Calendar, Event


def parse_checklist_log(file_content):
    pattern = r'-\s*\[([ xX])\]\s*\*\*(\d{4}-\d{2}-\d{2}\s*\([^)]+\))\*\*(.*?)(?=\n\s*-\s*\[[ xX]\]\s*\*\*\d{4}-\d{2}-\d{2}|\Z)'

    matches = re.findall(
        pattern,
        file_content,
        re.DOTALL
    )

    entries = []

    for status, date_str, content in matches:
        raw_content = content.strip()
        
        # Split out the title line and remove any leading colons or hyphens left over from markdown
        lines = raw_content.split("\n")
        title = lines[0].strip().lstrip(":- ")

        # Rebuild the remaining content without the first line and clean up leading punctuation
        body_lines = [line.strip().lstrip(":- ") for line in lines[1:] if line.strip()]
        cleaned_content = "\n".join(body_lines)

        # Convert markdown bullets into calendar-friendly lines
        cleaned_content = re.sub(
            r'\n\s*[-*]\s+',
            '\n• ',
            cleaned_content
        )

        # Extract just the YYYY-MM-DD date part from the date_str group
        clean_date = re.search(r'\d{4}-\d{2}-\d{2}', date_str).group(0)

        entries.append({
            "date": clean_date,
            "title": title,
            "content": cleaned_content if cleaned_content else title,
            "completed": status.upper() == "X"
        })

    return entries


def is_out_of_office(entry):
    """
    Detect entries that should be marked as Outlook Out Of Office.
    Checks both the heading and the work log content.
    """

    text = (
        entry.get("title", "")
        + " "
        + entry.get("content", "")
    ).lower()

    keywords = [
        "day off",
        "holiday",
        "public holiday",
        "annual leave",
        "vacation",
        "pto"
    ]

    return any(
        keyword in text
        for keyword in keywords
    )


def create_ics(entries):

    cal = Calendar()

    cal.add(
        'prodid',
        '-//Worklog Automation//mxp//'
    )

    cal.add(
        'version',
        '2.0'
    )

    for entry in entries:

        event = Event()

        if is_out_of_office(entry):

            event.add(
                'summary',
                "Out of Office"
            )

            event.add(
                'X-MICROSOFT-CDO-BUSYSTATUS',
                'OOF',
                parameters={'VALUE': 'TEXT'}
            )

            event.add(
                'X-MICROSOFT-CDO-INTENDEDSTATUS',
                'OOF',
                parameters={'VALUE': 'TEXT'}
            )

        else:

            # Normal work logs are regular calendar events
            # and do not affect availability
            event.add(
                'summary',
                f"Work Log: {entry['date']}"
            )

        description = entry["content"]

        # Ensure bullets render correctly in Outlook
        description = re.sub(
            r'\s*\*\s+',
            '\r\n• ',
            description
        )

        description = re.sub(
            r'\s*-\s+',
            '\r\n• ',
            description
        )

        event.add(
            'description',
            description
        )

        d = datetime.strptime(
            entry["date"],
            "%Y-%m-%d"
        ).date()

        event.add(
            'dtstart',
            d
        )

        event.add(
            'dtend',
            d
        )

        event.add(
            'uid',
            f"worklog-{entry['date']}@worklog.automation"
        )

        cal.add_component(event)

    return cal.to_ical()


def send_email(ics_content, file_name):

    connection_string = os.environ[
        "AZURE_COMMUNICATION_CONNECTION_STRING"
    ]

    client = EmailClient.from_connection_string(
        connection_string
    )

    sender_address = os.environ["SENDER_EMAIL"]
    recipient_address = os.environ["RECIPIENT_EMAIL"]

    message = {
        "senderAddress": sender_address,
        "recipients": {
            "to": [
                {
                    "address": recipient_address
                }
            ]
        },
        "content": {
            "subject": f"Work Log ICS: {file_name}",
            "plainText": (
                "Attached are your calendar records "
                "parsed from your monthly markdown checklist."
            )
        },
        "attachments": [
            {
                "name": f"{file_name}.ics",
                "contentType": "text/calendar",
                "contentInBase64": ics_content
            }
        ]
    }

    poller = client.begin_send(message)

    print(
        f"Email send result: {poller.result()}"
    )


if __name__ == "__main__":

    file_path = sys.argv[1]

    file_name = os.path.splitext(
        os.path.basename(file_path)
    )[0]

    with open(
        file_path,
        "r",
        encoding="utf-8"
    ) as f:
        content = f.read()

    entries = parse_checklist_log(content)

    if not entries:
        print(
            "No valid checklist dates found in the markdown file."
        )
        sys.exit(0)

    ics_bytes = create_ics(entries)

    ics_base64 = base64.b64encode(
        ics_bytes
    ).decode("utf-8")

    send_email(
        ics_base64,
        file_name
    )