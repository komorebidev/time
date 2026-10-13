# Work Log Reporting

A Python Flask app that generates  PDF work logs from Markdown files with optional weekly summaries using Cloudflare Workers AI.

## Run command

```powershell
$ip = (Invoke-RestMethod -Uri "https://api.ipify.org").Trim()

Write-Host "Allowing public IP: $ip"

az deployment sub create `
  --name "time-deployment" `
  --location japaneast `
  --template-file main.bicep `
  --parameters location=japaneast `
  --parameters allowedIp="['$($ip)/32']" `
  --no-wait
```

## Getting azure app url

```powershell
$subId = az account show --query id --output tsv
$suffix = $subId.Substring($subId.Length - 4)
$rg = "time-$suffix"
$app = "time-$suffix"

az webapp show `
  --resource-group $rg `
  --name $app `
  --query "defaultHostName" `
  --output tsv
  ```

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

* Set these App Service environment variables in the portal:
- GITHUB_TOKEN
- CLOUDFLARE_ACCOUNT_ID
- CLOUDFLARE_API_TOKEN

The bicep already adds the variables but they need to be filled in from the portal.

### Steps in Azure Portal
1. Open the Azure Portal.
2. Navigate to your resource group, worklog-reporting-6411 (replace 6411 with your actual suffix).
3. Open your App Service (worklog-flask-6411).
4. Under Settings, select Environment variables.
5. In App settings, add these three entries:

# Local Testing

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```