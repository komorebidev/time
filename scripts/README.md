# Sync Module

Push work logs to calendar as ics invites using Azure resources to send the email

# Features

* Converts markdown to Outlook ICS format
* Supports multiple markdown file changes at once
* Does not create duplicate events
* If keywords like "Day off" are detected, it sets the day as out of office
* Office location support (if wfh or remote is detected for that day)

## Resources

### Azure Communication Services

* SMTP service

### Email Communication Services

* Sending the ics invite

### Azure Free subdomain

* Sender address

### Github Actions

* Powering the sync service

# Run command

* Deploying the Azure resources

```powershell
az login
```
* The --location argument in the deployment commands are only for metadata of the logs
* Locations defined in the code remain as is

* What-if

```powershell
az deployment sub what-if --location koreacentral --template-file main.bicep --query properties.outputs.communicationServiceSecret.value -o tsv
```

* Production

```powershell
az deployment sub create --location koreacentral --template-file main.bicep --query properties.outputs.communicationServiceSecret.value -o tsv
```

* Use the output string and paste into Github repository secrets as AZURE_COMMUNICATION_CONNECTION_STRING

## Setting up default Azure sender email

* Need to provision this to get the value for Github Actions
* For setting the default sender email

```powershell
az communication email domain sender-username create --email-service-name "worklog-emailService" --domain-name "AzureManagedDomain" --resource-group "worklog-rg" --name "donotreply" --username "donotreply" --display-name "Calendar Sync"
```

* For printing the default sender email value (look for fromSenderDomain in the output)

```powershell
az communication email domain show --email-service-name "worklog-emailService" --resource-group "worklog-rg" --name "AzureManagedDomain"
```

* DoNotReply+fromSenderDomain

* i.e. DoNotReply@035352be-63e4-4683-9aec-58f2ca97a47a.azurecomm.net

### Warning

* Destroying Azure resources and rebuilding will generate new communicationService secret

* This secret is used in Github Actions to send ics invites from Azure

* If rebuilding resources, update the key in Github Actions

## Filling in the repository secrets

* Navigate to Settings > Secrets and variables > Actions
* Click New repository secret and add the following three secrets one by one:
* AZURE_COMMUNICATION_CONNECTION_STRING
* SENDER_EMAIL
* RECIPIENT_EMAIL

# Cleanup

```powershell
az group delete --name ***
az group delete --name worklog-rg
```
# Troubleshooting

* To check email delivery logs, need to register this resource in Azure
* Easiest to set the recipient as personal email because Azure subdomains can get blocked by corporate settings
* Check spam too

```powershell
az provider register --namespace Microsoft.Insights
```