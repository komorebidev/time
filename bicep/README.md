# Sync Module

Push work logs to calendar as ics invites using Azure resources

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

* What-if

```powershell
az deployment sub what-if --location koreacentral --template-file main.bicep --query properties.outputs.communicationServiceSecret.value -o tsv
```

* Production

```powershell
az deployment sub create --location koreacentral --template-file main.bicep --query properties.outputs.communicationServiceSecret.value -o tsv
```

* Use the output string and paste into Github repository secrets as AZURE_COMMUNICATION_CONNECTION_STRING

### Warning

* Destroying Azure resources and rebuilding will generate new communicationService secret

* This secret is used in Github Actions to send ics invites from Azure

* If rebuilding resources, update the key in Github Actions