//declarations

param location string
param resourcePrefix string

// Create the Email Communication Service container
resource emailService 'Microsoft.Communication/emailServices@2023-03-31' = {
  name: '${resourcePrefix}-emailService'
  location: location
  properties: {
    dataLocation: 'korea'
  }
}

// Provision an Azure-managed domain under the email service
resource azureManagedDomain 'Microsoft.Communication/emailServices/domains@2023-03-31' = {
  parent: emailService
  name: 'AzureManagedDomain'
 // name: '${resourcePrefix}-azureDomain' //custom names are unsupported so comment out
  
 location: location
  properties: {
    domainManagement: 'AzureManaged'
  }
}

// Create the core Communication Service and link it to the managed email domain
resource communicationService 'Microsoft.Communication/communicationServices@2023-03-31' = {
  name: '${resourcePrefix}-communicationService'
  location: location
  properties: {
    dataLocation: 'korea'
    linkedDomains: [
      azureManagedDomain.id
    ]
  }
}


#disable-next-line outputs-should-not-contain-secrets //disables lint check for secrets
output communicationserviceSecret string = communicationService.listKeys().primaryConnectionString
// for getting the communicationService secret for github actions ↑
