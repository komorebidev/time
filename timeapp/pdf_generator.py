from pathlib import Path
import re
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


BASE_DIR = Path(__file__).resolve().parent
FONT_PATH = BASE_DIR / "NotoSansJP.ttf"

if not FONT_PATH.is_file():
    raise FileNotFoundError(f"Japanese font not found: {FONT_PATH}")

pdfmetrics.registerFont(TTFont("NotoSansJP", str(FONT_PATH)))


def clean_text(value):
    """Clean text for PDF output."""
    if value is None:
        return ""

    text = str(value)

    # Remove supplementary Unicode characters, including most emoji.
    text = "".join(char for char in text if ord(char) <= 0xFFFF)

    # Remove control characters except common whitespace.
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)

    return text.strip()


def escape_text(value):
    """Escape text for ReportLab Paragraph markup."""
    return escape(clean_text(value)).replace("\n", "<br/>")


def generate_weekly_pdf(
    entries,
    start_date,
    end_date,
    summary="",
    output_path="weekly_report.pdf",
):
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=f"Weekly Work Report {start_date} - {end_date}",
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Title"],
        fontName="NotoSansJP",
        fontSize=18,
        leading=24,
        alignment=TA_CENTER,
        spaceAfter=6 * mm,
    )

    subtitle_style = ParagraphStyle(
        "ReportSubtitle",
        parent=styles["Normal"],
        fontName="NotoSansJP",
        fontSize=10,
        leading=15,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#555555"),
        spaceAfter=8 * mm,
    )

    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading2"],
        fontName="NotoSansJP",
        fontSize=12,
        leading=17,
        spaceBefore=4 * mm,
        spaceAfter=2 * mm,
    )

    body_style = ParagraphStyle(
        "ReportBody",
        parent=styles["BodyText"],
        fontName="NotoSansJP",
        fontSize=9,
        leading=14,
        alignment=TA_LEFT,
        spaceAfter=2 * mm,
        wordWrap="CJK",
    )

    small_style = ParagraphStyle(
        "ReportSmall",
        parent=styles["Normal"],
        fontName="NotoSansJP",
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#555555"),
    )

    story = [
        Paragraph("Weekly Work Report", title_style),
        Paragraph(
            f"{escape_text(start_date)} - {escape_text(end_date)}",
            subtitle_style,
        ),
    ]

    if summary:
        story.append(Paragraph("Weekly Summary", heading_style))
        story.append(Paragraph(escape_text(summary), body_style))
        story.append(Paragraph("AI-generated summary", small_style))
        story.append(Spacer(1, 4 * mm))

    story.append(Paragraph("Work Log", heading_style))

    if not entries:
        story.append(
            Paragraph("No work-log entries for this week.", body_style)
        )
    else:
        table_data = [
            [
                Paragraph("Date", small_style),
                Paragraph("Work Details", small_style),
            ]
        ]

        for entry in entries:
            if isinstance(entry, dict):
                date_value = entry.get("date") or entry.get("day") or ""

                # The work-log parser stores tasks in a list named "tasks".
                tasks = entry.get("tasks")

                if isinstance(tasks, list):
                    task_value = "\n".join(
                        str(task) for task in tasks if task
                    )
                else:
                    # Retain support for older or differently shaped entries.
                    task_value = (
                        entry.get("task")
                        or entry.get("text")
                        or entry.get("content")
                        or ""
                    )

                metadata = entry.get("metadata")

                if metadata:
                    if isinstance(metadata, dict):
                        metadata_text = " | ".join(
                            f"{key}: {value}"
                            for key, value in metadata.items()
                        )
                    else:
                        metadata_text = str(metadata)

                    if metadata_text:
                        task_value = (
                            f"{task_value}\n{metadata_text}"
                            if task_value
                            else metadata_text
                        )

            else:
                date_value = getattr(entry, "date", "")
                task_value = (
                    getattr(entry, "task", None)
                    or getattr(entry, "text", None)
                    or str(entry)
                )

            table_data.append(
                [
                    Paragraph(escape_text(date_value), small_style),
                    Paragraph(escape_text(task_value), body_style),
                ]
            )

        table = Table(
            table_data,
            colWidths=[30 * mm, 144 * mm],
            repeatRows=1,
            hAlign="LEFT",
        )

        table.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.HexColor("#E8EEF5"),
                    ),
                    (
                        "TEXTCOLOR",
                        (0, 0),
                        (-1, 0),
                        colors.HexColor("#222222"),
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.35,
                        colors.HexColor("#CCCCCC"),
                    ),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )

        story.append(table)

    doc.build(story)
    return str(output_path)
