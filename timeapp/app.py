from datetime import date, timedelta
from io import BytesIO

from flask import Flask, render_template, request, send_file

from cloudflare_ai import summarize_worklog
from github_client import get_monthly_log
from worklog_parser import parse_worklog, filter_entries
from pdf_generator import generate_weekly_pdf

app = Flask(__name__)


def get_week_entries(start_date: date, end_date: date) -> list[dict]:
    """Retrieve and combine monthly logs covering the selected week."""

    months = {
        (start_date.year, start_date.month),
        (end_date.year, end_date.month),
    }

    all_entries = []

    for year, month in sorted(months):
        content = get_monthly_log(year, month)
        all_entries.extend(parse_worklog(content))

    return filter_entries(all_entries, start_date, end_date)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/download", methods=["POST"])
def download():
    week_start = request.form.get("week_start", "")

    try:
        start_date = date.fromisoformat(week_start)
    except ValueError:
        return "Please select a valid date.", 400

    # Always use Monday as the start of the selected week.
    start_date -= timedelta(days=start_date.weekday())
    end_date = start_date + timedelta(days=6)

    try:
        # Retrieve the selected week's work-log entries.
        entries = get_week_entries(start_date, end_date)

        # Generate an English summary using Cloudflare Workers AI.
        summary = summarize_worklog(entries)

        # Generate the PDF, including the AI summary.
        pdf_buffer = BytesIO()

        generate_weekly_pdf(
            entries,
            start_date,
            end_date,
            summary,
            "weekly_report.pdf",
        )

        with open("weekly_report.pdf", "rb") as pdf_file:
            pdf_buffer.write(pdf_file.read())

        pdf_buffer.seek(0)

        filename = (
            f"work-report-{start_date:%Y-%m-%d}"
            f"-to-{end_date:%Y-%m-%d}.pdf"
        )

        return send_file(
            pdf_buffer,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename,
        )

    except Exception as error:
        app.logger.exception("Failed to generate work report")
        return f"Could not generate the report: {error}", 500


if __name__ == "__main__":
    app.run(debug=True)
