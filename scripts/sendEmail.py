import os
import sys
import re
import base64
from datetime import datetime
from azure.communication.email import EmailClient
from icalendar import Calendar, Event

def parse_checklist_log(file_content):
    pattern = r'-\s*\[([ xX])\]\s*\*\*(\d{4}-\d{2}-\d{2})\s*\([^)]+\)\*\*:\s*(.*?)(?=\n\s*-\s*\[[ xX]\]\s*\*\*\d{4}-\d{2}-\d{2}|\Z)'
    matches = re.findall(pattern, file_content, re.DOTALL)
    
    entries = []
    for status, date_str, content in matches:
        cleaned_content = content.strip()
        entries.append({
            "date": date_str.strip(),
            "content": cleaned_content if cleaned_content else "Work log entry",
            "completed": status.upper() == 'X'
        })
    return entries

def create_ics(entries):
    cal = Calendar()
    cal.add('prodid', '-//Worklog Automation//mxp//')
    cal.add('version', '2.0')

    for entry in entries:
        event = Event()
        event.add('summary', f"Work Log: {entry['date']}")
        event.add('description', entry['content'])
        
        d = datetime.strptime(entry['date'], '%Y-%m-%d').date()
        event.add('dtstart', d)
        event.add('dtend', d)
        event.add('uid', f"worklog-{entry['date']}@worklog.automation")
        
        cal.add_component(event)
        
    return cal.to_ical()

def send_email(ics_content, file_name):
    connection_string = os.environ["AZURE_COMMUNICATION_CONNECTION_STRING"]
    client = EmailClient.from_connection_string(connection_string)

    sender_address = os.environ["SENDER_EMAIL"] 
    recipient_address = os.environ["RECIPIENT_EMAIL"]

    message = {
        "senderAddress": sender_address,
        "recipients": {
            "to": [{"address": recipient_address}]
        },
        "content": {
            "subject": f"Work Log ICS: {file_name}",
            "plainText": "Attached are your calendar records parsed from your monthly markdown checklist.",
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
    print(f"Email send result: {poller.result()}")

if __name__ == "__main__":
    file_path = sys.argv[1]
    file_name = os.path.splitext(os.path.basename(file_path))[0]

    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    entries = parse_checklist_log(content)
    
    if not entries:
        print("No valid checklist dates found in the markdown file.")
        sys.exit(0)

    ics_bytes = create_ics(entries)
    ics_base64 = base64.b64encode(ics_bytes).decode('utf-8')

    send_email(ics_base64, file_name)