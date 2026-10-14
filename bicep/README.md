# Work Log Reporting

A Python Flask app that generates  PDF work logs from Markdown files with optional weekly summaries using Cloudflare Workers AI.

* Note that the IP restriction means can only deploy from allowed source IP

## Azure App Service – Free Tier Behaviour

The Flask weekly work report app is hosted on **Azure App Service Free tier (F1)**.

- **No scale-to-zero:** App Service does not scale the application to zero instances like Azure Container Apps can.
- **Idle behaviour:** The application may become idle when there is no traffic. The first request after inactivity may take longer to respond.
- **Shared infrastructure:** The Free tier uses shared compute resources and has limited CPU time and resource quotas.
- **Always On unavailable:** The Always On setting is not supported on the Free tier, so the application cannot be configured to stay continuously active through that setting.
- **Scaling limitations:** The Free tier does not support scaling out to multiple instances.

This tier is suitable for a low-traffic personal work report application, provided its resource limits are sufficient.

### Verify the App Service Plan

```powershell
az appservice plan list --query "[].{Name:name, ResourceGroup:resourceGroup, Location:location, Tier:sku.tier, SKU:sku.name, Capacity:sku.capacity}" -o table
```

## Run command

```powershell
$ip = (Invoke-RestMethod -Uri "https://api.ipify.org").Trim()

Write-Host "Allowing public IP: $ip"

az deployment sub create `
  --name "time-deployment" `
  --location japaneast `
  --template-file main.bicep `
  --parameters location=japaneast `
  --parameters allowedIp='['$($ip)/32']' `
  --no-wait
```

## Uploading workload to azure app (index.html and etc)

* This is the actual contents of the azure app which will be hosted
* Until now, it is only infra

```powershell
$subId = az account show --query id --output tsv
$suffix = $subId.Substring($subId.Length - 4)
$rg = "time-$suffix"
$app = "time-$suffix"

az webapp deploy `
  --resource-group $rg `
  --name $app `
  --src-path ".\time-app.zip" `
  --type zip
```

## Fix: ZIP Deployment on Azure App Service (Linux)

**Problem:** The application deployed, but Flask couldn't find `templates/index.html`. The ZIP archive contained the path `templates\\index.html` with Windows-style backslashes, which caused the Linux deployment to fail to locate the template.

**Solution:** Create the ZIP using Python's `zipfile` module and explicitly specify archive paths with forward slashes.

Run from the project directory:

```powershell
Set-Location "$HOME\Downloads\time-app"

@'
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

root = Path.cwd()
output = root.parent / "time-app.zip"

files = [
    "app.py",
    "cloudflare_ai.py",
    "github_client.py",
    "pdf_generator.py",
    "worklog_parser.py",
    "requirements.txt",
    "NotoSansJP.ttf",
    "templates/index.html",
]

with ZipFile(output, "w", ZIP_DEFLATED) as z:
    for name in files:
        z.write(root / name, arcname=name)

print(f"Created {output}")
'@ | python -
```

Verify the archive contains the correct path:

```powershell
tar -tf "$HOME\Downloads\time-app.zip" |
    Select-String "templates"
```

Expected output:

```text
templates/index.html
```

Deploy the ZIP:

```powershell
az webapp deploy `
    --resource-group $rg `
    --name $app `
    --src-path "$HOME\Downloads\time-app.zip" `
    --type zip
```

**Key lesson:** When deploying Python applications to Linux App Service, ensure ZIP entry paths use forward slashes (`/`) and that application files are at the archive root, without an extra project folder.

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