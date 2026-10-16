import os
import sys
import re
import subprocess
from datetime import datetime, timedelta

from icalendar import Calendar, Event


# ======================================================================
# Find Markdown files (Git-aware incremental scan or full scan)
# ======================================================================

def find_markdown_files(path, full_refresh=False):
    markdown_files = []

    # Full scan if requested or if the path is not a Git repository.
    if full_refresh or not os.path.exists(os.path.join(path, ".git")):
        print("Performing full scan of Markdown files...")

        if os.path.isfile(path):
            if path.endswith(".md"):
                markdown_files.append(path)
        elif os.path.isdir(path):
            for root, dirs, files in os.walk(path):
                for filename in files:
                    if filename.endswith(".md"):
                        markdown_files.append(os.path.join(root, filename))

        return markdown_files

    # Incremental mode: identify Markdown files changed in the latest commit.
    try:
        print("Detecting changed Markdown files via Git...")
        result = subprocess.run(
            ["git", "diff", "--name-only", "HEAD~1", "HEAD"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True,
        )
        changed_files = result.stdout.splitlines()

        result_untracked = subprocess.run(
            ["git", "ls-files", "--others", "--exclude-standard"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True,
        )
        untracked_files = result_untracked.stdout.splitlines()

        all_changed = set(changed_files + untracked_files)
        for filename in all_changed:
            if filename.endswith(".md") and os.path.exists(filename):
                markdown_files.append(filename)

        if not markdown_files:
            print("Git reports no Markdown files changed in this commit.")

    except Exception as exc:
        print(f"Git diff failed ({exc}), falling back to full scan...")
        return find_markdown_files(path, full_refresh=True)

    return markdown_files


# ======================================================================
# Markdown heading log parser
# ======================================================================

def parse_checklist_log(file_content):
    # Matches headings such as:
    # ## **2026-01-01 (Thursday)**
    # and captures the heading and its content up to the next ## heading.
    pattern = re.compile(
        r"^##\s*(\*{0,2}\d{4}-\d{2}-\d{2}\s*\([^)]*\)\*{0,2}.*?)(?=^##\s|\Z)",
        re.DOTALL | re.MULTILINE,
    )

    entries = []
    for match in pattern.finditer(file_content):
        raw_content = match.group(1).strip()
        lines = raw_content.splitlines()
        if not lines:
            continue

        title = lines[0].strip()
        body_lines = [line.strip().lstrip(":- ") for line in lines[1:] if line.strip()]
        cleaned_content = "\n".join(body_lines)
        cleaned_content = re.sub(r"\n\s*[-*]\s+", "\n• ", cleaned_content)

        date_match = re.search(r"\d{4}-\d{2}-\d{2}", title)
        if not date_match:
            continue

        entries.append(
            {
                "date": date_match.group(0),
                "title": title,
                "content": cleaned_content if cleaned_content else "",
            }
        )

    return entries


# ======================================================================
# Remove Markdown syntax for iCloud titles
# ======================================================================

def clean_markdown(text):
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"\*(.*?)\*", r"\1", text)
    text = text.replace("`", "")
    return text.strip()


# ======================================================================
# Event detection
# ======================================================================

def is_public_holiday(entry):
    text = entry.get("title", "") + " " + entry.get("content", "")
    return "🎌" in text or "public holiday" in text.lower()


def is_out_of_office(entry):
    title = entry.get("title", "")
    text = title + " " + entry.get("content", "")

    patterns = [
        r"\bday[\s-]?off\b",
        r"\b(am|pm)[\s-]?off\b",
        r"\btaking\s+leave\b",
        r"\bpaid\s+leave\b",
        r"\bannual\s+leave\b",
        r"\bpersonal\s+leave\b",
        r"\bsick\s+leave\b",
        r"\bon\s+leave\b",
        r"\bleave\b",
        r"\btime\s+off\b",
        r"\bvacation\b",
        r"\bpto\b",
        r"\bholiday\b",
    ]
    match_general = any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)
    match_header_off = re.search(r"\boff\b", title, re.IGNORECASE) is not None

    return is_public_holiday(entry) or match_general or match_header_off


def is_wfh(entry):
    text = (entry.get("title", "") + " " + entry.get("content", "")).lower()
    keywords = ["wfh", "remote", "working from home", "work from home"]
    return any(keyword in text for keyword in keywords)


def is_customer_visit(entry):
    title = entry.get("title", "")