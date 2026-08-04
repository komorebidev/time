import os
import sys
import re
import base64

from datetime import datetime, timedelta

from azure.communication.email import EmailClient
from icalendar import Calendar, Event



# ======================================================================
# Find markdown files recursively
# ======================================================================

def find_markdown_files(path):

    markdown_files = []


    if os.path.isfile(path):

        if path.endswith(".md"):

            markdown_files.append(path)



    elif os.path.isdir(path):

        for root, dirs, files in os.walk(path):

            for file in files:

                if file.endswith(".md"):

                    markdown_files.append(
                        os.path.join(
                            root,
                            file
                        )
                    )


    return markdown_files





# ======================================================================
# Markdown checklist parser
# ======================================================================

def parse_checklist_log(file_content):

    pattern = (
        r'-\s*\[([ xX])\]\s*'
        r'\*\*(\d{4}-\d{2}-\d{2}\s*\([^)]+\))\*\*'
        r'(.*?)(?=\n\s*-\s*\[[ xX]\]\s*\*\*\d{4}-\d{2}-\d{2}|\Z)'
    )


    matches = re.findall(
        pattern,
        file_content,
        re.DOTALL
    )


    entries = []


    for status, date_str, content in matches:


        raw_content = content.strip()


        lines = raw_content.split("\n")


        title = (
            lines[0]
            .strip()
            .lstrip(":- ")
        )


        body_lines = [

            line.strip().lstrip(":- ")

            for line in lines[1:]

            if line.strip()

        ]


        cleaned_content = "\n".join(
            body_lines
        )


        cleaned_content = re.sub(
            r'\n\s*[-*]\s+',
            '\n• ',
            cleaned_content
        )


        clean_date = re.search(
            r'\d{4}-\d{2}-\d{2}',
            date_str
        ).group(0)



        entries.append(
            {
                "date": clean_date,

                "title": title,

                "content": (
                    cleaned_content
                    if cleaned_content
                    else title
                ),

                "completed": (
                    status.upper() == "X"
                )
            }
        )


    return entries





# ======================================================================
# Remove markdown syntax for iCloud titles
# ======================================================================

def clean_markdown(text):


    # Remove bold formatting
    text = re.sub(
        r"\*\*(.*?)\*\*",
        r"\1",
        text
    )


    # Remove italic formatting
    text = re.sub(
        r"\*(.*?)\*",
        r"\1",
        text
    )


    # Remove inline code markers
    text = text.replace(
        "`",
        ""
    )


    return text.strip()





# ======================================================================
# Event detection
# ======================================================================

def is_out_of_office(entry):

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





def is_wfh(entry):

    text = (

        entry.get("title", "")

        + " "

        + entry.get("content", "")

    ).lower()


    keywords = [

        "wfh",

        "remote",

        "working from home",

        "work from home"

    ]


    return any(

        keyword in text

        for keyword in keywords

    )





def is_customer_visit(entry):

    title = entry.get(
        "title",
        ""
    )


    return (

        re.search(
            r"\bon-?site\b",
            title,
            re.IGNORECASE
        )

        is not None

        or "onsite" in title.lower()

    )





# ======================================================================
# iCloud calendar summary mapping
# ======================================================================

def get_icloud_summary(entry):


    title = clean_markdown(
        entry.get(
            "title",
            ""
        )
    )


    content = clean_markdown(
        entry.get(
            "content",
            ""
        )
    )


    combined = (
        title
        + " "
        + content
    ).lower()



    # Preserve public holiday names
    if "public holiday" in combined:

        return title



    if is_out_of_office(entry):

        return "🌴 休み"



    if is_customer_visit(entry):

        return "🏢 客先"



    if is_wfh(entry):

        return "🏠 在宅"



    return "🏢 出社"

# ======================================================================
# Outlook ICS generation
# ======================================================================
#
# Used for corporate Outlook calendar.
#
# Includes:
#   - Full work descriptions
#   - Customer information
#   - WFH / customer visit locations
#   - Outlook Out Of Office status
#
# ======================================================================


def create_outlook_ics(entries):

    cal = Calendar()


    cal.add(
        "prodid",
        "-//Worklog Automation//mxp//"
    )


    cal.add(
        "version",
        "2.0"
    )



    for entry in entries:


        event = Event()



        if is_out_of_office(entry):


            event.add(
                "summary",
                "Out of Office"
            )


            # Outlook OOF availability status

            event.add(
                "X-MICROSOFT-CDO-BUSYSTATUS",
                "OOF",
                parameters={
                    "VALUE": "TEXT"
                }
            )


            event.add(
                "X-MICROSOFT-CDO-INTENDEDSTATUS",
                "OOF",
                parameters={
                    "VALUE": "TEXT"
                }
            )



        else:


            event.add(
                "summary",
                f"Work Log: {entry['date']}"
            )



            if is_customer_visit(entry):

                event.add(
                    "location",
                    "Customer Visit"
                )


            elif is_wfh(entry):

                event.add(
                    "location",
                    "WFH"
                )



        # Outlook receives full descriptions

        description = entry["content"]


        description = re.sub(
            r"\s*\*\s+",
            "\r\n• ",
            description
        )


        description = re.sub(
            r"\s*-\s+",
            "\r\n• ",
            description
        )


        event.add(
            "description",
            description
        )



        d = datetime.strptime(
            entry["date"],
            "%Y-%m-%d"
        ).date()



        event.add(
            "dtstart",
            d
        )


        event.add(
            "dtend",
            d + timedelta(days=1)
        )



        event.add(
            "uid",
            f"outlook-{entry['date']}@worklog"
        )



        cal.add_component(
            event
        )



    return cal.to_ical()







# ======================================================================
# iCloud ICS generation
# ======================================================================
#
# Used for Apple Calendar subscription.
#
# Output:
#
#     docs/time.ics
#
#
# Contains:
#   - Summary
#   - Date
#
#
# Does NOT contain:
#   - Work descriptions
#   - Customer notes
#   - Task details
#
# ======================================================================


def create_icloud_ics(entries):

    cal = Calendar()



    cal.add(
        "prodid",
        "-//Worklog Automation iCloud//mxp//"
    )


    cal.add(
        "version",
        "2.0"
    )



    # iCloud subscribed calendar name

    cal.add(
        "X-WR-CALNAME",
        "仕事"
    )


    cal.add(
        "X-WR-CALDESC",
        "自動生成された勤務予定"
    )



    for entry in entries:


        event = Event()



        event.add(
            "summary",
            get_icloud_summary(entry)
        )



        d = datetime.strptime(
            entry["date"],
            "%Y-%m-%d"
        ).date()



        event.add(
            "dtstart",
            d
        )


        event.add(
            "dtend",
            d + timedelta(days=1)
        )



        event.add(
            "uid",
            f"icloud-{entry['date']}@worklog"
        )



        cal.add_component(
            event
        )



    return cal.to_ical()

# ======================================================================
# Outlook ICS email sending
# ======================================================================
#
# Sends the corporate Outlook calendar.
#
# Email:
#   Subject: Work Log ICS: time
#
# Attachment:
#   time.ics
#
# ======================================================================


def send_email(ics_content, file_name):


    connection_string = os.environ[
        "AZURE_COMMUNICATION_CONNECTION_STRING"
    ]


    client = EmailClient.from_connection_string(
        connection_string
    )


    sender_address = os.environ[
        "SENDER_EMAIL"
    ]


    recipient_address = os.environ[
        "RECIPIENT_EMAIL"
    ]



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

            "subject": (
                f"Work Log ICS: {file_name}"
            ),


            "plainText": (
                "Attached are your work calendar records "
                "generated from markdown files."
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



    poller = client.begin_send(
        message
    )


    print(
        f"Email send result: {poller.result()}"
    )







# ======================================================================
# Main execution
# ======================================================================
#
# Run:
#
# python scripts/sendEmail.py .
#
#
# Scans all markdown files:
#
# 2026/
# 2027/
# etc.
#
#
# Generates:
#
# Outlook:
#   Email -> time.ics
#
# iCloud:
#   docs/time.ics
#
# ======================================================================


if __name__ == "__main__":


    if len(sys.argv) < 2:

        print(
            "Usage: python sendEmail.py <file-or-folder>"
        )

        sys.exit(1)



    input_path = sys.argv[1]



    markdown_files = find_markdown_files(
        input_path
    )



    if not markdown_files:

        print(
            "No markdown files found."
        )

        sys.exit(0)



    print(
        "Markdown files discovered:"
    )


    for file in markdown_files:

        print(
            f" - {file}"
        )



    # --------------------------------------------------------------
    # Parse all markdown files
    # --------------------------------------------------------------

    all_entries = []



    for file in markdown_files:


        with open(
            file,
            "r",
            encoding="utf-8"
        ) as f:

            content = f.read()



        entries = parse_checklist_log(
            content
        )


        all_entries.extend(
            entries
        )



    if not all_entries:

        print(
            "No calendar entries found."
        )

        sys.exit(0)



    # --------------------------------------------------------------
    # Sort events by date
    # --------------------------------------------------------------

    all_entries.sort(
        key=lambda x: x["date"]
    )



    print(
        f"Total calendar entries: {len(all_entries)}"
    )



    # --------------------------------------------------------------
    # Generate Outlook ICS
    # --------------------------------------------------------------

    outlook_ics = create_outlook_ics(
        all_entries
    )


    outlook_base64 = base64.b64encode(
        outlook_ics
    ).decode(
        "utf-8"
    )



    # Outlook email:
    #
    # Subject:
    #   Work Log ICS: time
    #
    # Attachment:
    #   time.ics

    send_email(
        outlook_base64,
        "time"
    )



    # --------------------------------------------------------------
    # Generate iCloud ICS
    # --------------------------------------------------------------

    icloud_ics = create_icloud_ics(
        all_entries
    )



    os.makedirs(
        "docs",
        exist_ok=True
    )



    with open(
        "docs/time.ics",
        "wb"
    ) as f:

        f.write(
            icloud_ics
        )



    print(
        "Created iCloud calendar feed: docs/time.ics"
    )