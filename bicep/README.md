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

* donotreply+fromSenderDomain

* i.e. donotreply@035352be-63e4-4683-9aec-58f2ca97a47a.azurecomm.net

### Warning

* Destroying Azure resources and rebuilding will generate new communicationService secret

* This secret is used in Github Actions to send ics invites from Azure

* If rebuilding resources, update the key in Github Actions

# Cleanup

```powershell
az group delete --name ***
az group delete --name worklog-rg
```