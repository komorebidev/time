// declarations

targetScope = 'subscription' //for deploying all resources under the one resource group
param location string = 'global' //because azure communication and email are global, not regional
param resourcePrefix string = 'worklog'
param rgLocation string = 'koreacentral' 

// Create the Resource Group
resource resourceGroup 'Microsoft.Resources/resourceGroups@2023-07-01' = {
  name: '${resourcePrefix}-rg'
  location: rgLocation
  tags: { 
    Environment: 'POC' 
    ManagedBy: 'Bicep'
  }
}

//for calling syncModule.bicep (modules needed because creating resource groups needs subscription level but the other modules are resource level)
module syncModule 'syncModule.bicep' = {
  name: '${resourcePrefix}-syncmoduleDeployment'
  scope: resourceGroup // This satisfies the resource group scope requirement
  params: {
    resourcePrefix: resourcePrefix
    location: location
  }
}

// Output the connection string so you can save it as a GitHub Secret

output communicationServiceSecret string = syncModule.outputs.communicationserviceSecret
