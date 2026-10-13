targetScope = 'resourceGroup'

@description('Azure region for the App Service Plan.')
param location string

@description('Suffix derived from the subscription ID.')
param suffix string

var servicePlanName = 'time-plan-${suffix}'

resource servicePlan 'Microsoft.Web/serverfarms@2024-04-01' = {
  name: servicePlanName
  location: location
  kind: 'linux'
  sku: {
    name: 'B1'
    tier: 'Basic'
    size: 'B1'
    capacity: 1
  }
  properties: {
    reserved: true
  }
}

output appServicePlanId string = servicePlan.id