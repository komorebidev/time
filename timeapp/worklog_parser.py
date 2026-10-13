import re
from datetime import date

HEADING_PATTERN = re.compile(
    r"^##\s+\*{0,2}"
    r"(?P<date>\d{4}-\d{2}-\d{2})"
    r"(?:\s+\([^)]+\))?"
    r"\*{0,2}\s*:?\s*(?P<meta>.*)$"
)


def parse_worklog(markdown: str) -> list[dict]:
    entries = []
    current = None

    for line in markdown.splitlines():
        match = HEADING_PATTERN.match(line.strip())

        if match:
            if current:
                entries.append(current)

            current = {
                "date": date.fromisoformat(match.group("date")),
                "metadata": match.group("meta").strip(),
                "tasks": [],
            }
            continue

        if current:
            bullet = re.match(r"^\s*[-*+]\s+(.+)$", line)

            if bullet:
                current["tasks"].append(bullet.group(1).strip())

    if current:
        entries.append(current)

    return entries


def filter_entries(
    entries: list[dict],
    start_date: date,
    end_date: date,
) -> list[dict]:
    return [
        entry for entry in entries
        if start_date <= entry["date"] <= end_date
    ]
