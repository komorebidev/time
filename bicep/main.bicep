targetScope = 'subscription'

@description('Azure region for the resource group and application.')
param location string = 'japaneast'

@description('Public IP addresses or CIDR ranges allowed to access the app.')
param allowedIp array

var suffix = substring(subscription().id, length(subscription().id) - 4, 4)
var resourceGroupName = 'time${suffix}'

// Create the resource group.
resource rg 'Microsoft.Resources/resourceGroups@2024-03-01' = {
  name: resourceGroupName
  location: location
}

// Create the App Service Plan.
module servicePlan './serviceplan.bicep' = {
  scope: rg
  name: 'serviceplan-deployment'
  params: {
    location: location
    suffix: suffix
  }
}

// Create the Flask web application.
module flask './flask.bicep' = {
  scope: rg
  name: 'flask-deployment'
  params: {
    location: location
    suffix: suffix
    appServicePlanId: servicePlan.outputs.appServicePlanId
    allowedIp: allowedIp
  }
}

output resourceGroupName string = rg.name
output webAppName string = flask.outputs.webAppName
output webAppUrl string = flask.outputs.webAppUrl