@maxLength(20)
@minLength(4)
param resourceBaseName string
param staticWebAppSku string

@description('Required when create Azure Bot service')
param botAadAppClientId string

@secure()
@description('Required by Bot Framework package in your bot project')
param botAadAppClientSecret string

param webAppSKU string

@maxLength(42)
param botDisplayName string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName

param staticWebAppName string = resourceBaseName
param location string = resourceGroup().location

param aadClientId string
@secure()
param aadClientSecret string
param cosmosConnectionString string
param cosmosDbName string
param cosmosContainerName string
param eventGridEndpoint string
param eventGridKey string

// Azure Static Web Apps that hosts your static web site
resource swa 'Microsoft.Web/staticSites@2022-09-01' = {
  name: staticWebAppName
  // SWA do not need location setting
  location: 'centralus'
  sku: {
    name: staticWebAppSku
    tier: staticWebAppSku
  }
  properties: {}
}


// Compute resources for your Web App
resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSKU
  }
}

// Web App that hosts your bot
resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1' // Run Azure APP Service from a package file
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~18' // Set NodeJS version to 18.x for your site
        }
        {
          name: 'RUNNING_ON_AZURE'
          value: '1'
        }
        {
          name: 'BOT_ID'
          value: botAadAppClientId
        }
        {
          name: 'BOT_PASSWORD'
          value: botAadAppClientSecret
        }
        {
          name: 'AAD_APP_CLIENT_ID'
          value: aadClientId
        }
        {
          name: 'AAD_APP_CLIENT_SECRET'
          value: aadClientSecret
        }
        {
          name: 'TEAMS_APP_TENANT_ID'
          value: tenant().tenantId
        }
        {
          name: 'COSMOS_CONN_STRING'
          value: cosmosConnectionString
        }
        {
          name: 'COSMOS_DATABASE_NAME'
          value: cosmosDbName
        }
        {
          name: 'COSMOS_CONTAINER_NAME'
          value: cosmosContainerName
        }
        {
          name: 'EG_ENDPOINT'
          value: eventGridEndpoint
        }
        {
          name: 'EG_KEY'
          value: eventGridKey
        }
      ]
      ftpsState: 'FtpsOnly'
      cors: {
        allowedOrigins: [
          'https://${swa.properties.defaultHostname}'
        ]
      }
    }
  }
}

// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    botAadAppClientId: botAadAppClientId
    botAppDomain: webApp.properties.defaultHostName
    botDisplayName: botDisplayName
  }
}

var siteDomain = swa.properties.defaultHostname

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output AZURE_STATIC_WEB_APPS_RESOURCE_ID string = swa.id
output TAB_DOMAIN string = siteDomain
output TAB_HOSTNAME string = siteDomain
output TAB_ENDPOINT string = 'https://${siteDomain}'
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
