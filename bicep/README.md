# Work Log Reporting

A Python Flask app that generates  PDF work logs from Markdown files with optional weekly summaries using Cloudflare Workers AI.

## Architecture

- **Azure App Service** — hosts the Flask app.
- **GitHub** — stores work-log Markdown files.
- **Cloudflare Workers AI** — generates optional weekly summaries.
- **ReportLab** — generates PDF reports.

## Infrastructure

Deploy Azure resources using Bicep:

```powershell
az deployment sub create --name worklog-reporting-deployment --location japaneast --template-file infra/main.bicep
```
## Environment variables

* Set these App Service environment variables:
- GITHUB_TOKEN
- CLOUDFLARE_ACCOUNT_ID
- CLOUDFLARE_API_TOKEN

# Local Testing

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```