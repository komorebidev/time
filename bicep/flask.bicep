targetScope = 'resourceGroup'

@description('Azure region for the web app.')
param location string

@description('Suffix derived from the subscription ID.')
param suffix string

@description('Resource ID of the App Service Plan.')
param appServicePlanId string

@description('Public IPv4 addresses or CIDR ranges allowed to access the web app.')
param allowedIp array

var webAppName = 'time-${suffix}'

resource webApp 'Microsoft.Web/sites@2024-04-01' = {
  name: webAppName
  location: location
  kind: 'app,linux'
  properties: {
    httpsOnly: true
    serverFarmId: appServicePlanId
    reserved: true
    siteConfig: {
      linuxFxVersion: 'PYTHON|3.12'
      alwaysOn: false
      appCommandLine: 'gunicorn --bind=0.0.0.0:8000 --timeout 600 app:app'
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'

      // Allow only the configured public IP addresses or CIDR ranges.
      ipSecurityRestrictions: [
        for (cidr, i) in allowedIp: {
          name: 'Allow-IP-${i + 1}'
          ipAddress: cidr
          action: 'Allow'
          priority: 100 + i
          description: 'Allowed public IP range'
        }
      ]
      ipSecurityRestrictionsDefaultAction: 'Deny'

      // Apply the main site's IP restrictions to the SCM/Kudu endpoint too.
      scmIpSecurityRestrictionsUseMain: true

      appSettings: [
        {
          name: 'SCM_DO_BUILD_DURING_DEPLOYMENT'
          value: 'true'
        }
        {
          name: 'PYTHONUNBUFFERED'
          value: '1'
        }
        {
          name: 'GITHUB_TOKEN'
          value: ''
        }
        {
          name: 'CLOUDFLARE_ACCOUNT_ID'
          value: ''
        }
        {
          name: 'CLOUDFLARE_API_TOKEN'
          value: ''
        }
      ]
    }
  }
}

output webAppName string = webApp.name
output webAppUrl string = 'https://${webApp.properties.defaultHostName}'