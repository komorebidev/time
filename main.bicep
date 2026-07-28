// declarations
param location string = 'global' //because azure communication and email are global, not regional
param resourcePrefix string = 'worklog'

// Create the Email Communication Service container
resource emailService 'Microsoft.Communication/emailServices@2023-03-31' = {
  name: '${resourcePrefix}-emailService'
  location: location
  properties: {
    dataLocation: 'Korea Central'
  }
}

// Provision an Azure-managed domain under the email service
resource azureManagedDomain 'Microsoft.Communication/emailServices/domains@2023-03-31' = {
  parent: emailService
  name: '${resourcePrefix}-azureDomain'
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
    dataLocation: 'Korea Central'
    linkedDomains: [
      azureManagedDomain.id
    ]
  }
}

// Output the connection string so you can save it as a GitHub Secret

#disable-next-line outputs-should-not-contain-secrets //disables lint check for secrets
output communicationserviceconnectionString string = communicationService.listKeys().primaryConnectionString
