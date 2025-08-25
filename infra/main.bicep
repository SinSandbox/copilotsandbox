@description('Name of the Web App to create. Defaults to a unique name derived from the resource group.')
param webAppName string = 'copilotsandbox-${uniqueString(resourceGroup().id)}'
@description('Location for all resources')
param location string = resourceGroup().location
@description('SKU name for the App Service Plan')
param skuName string = 'B1'
@description('SKU tier')
param skuTier string = 'Basic'
@description('Capacity (number of instances)')
param skuCapacity int = 1

resource appServicePlan 'Microsoft.Web/serverfarms@2022-03-01' = {
  name: '${webAppName}-plan'
  location: location
  sku: {
    name: skuName
    tier: skuTier
    capacity: skuCapacity
  }
  properties: {
    // required for Linux workers
    reserved: true
  }
}

resource webApp 'Microsoft.Web/sites@2022-03-01' = {
  name: webAppName
  location: location
  kind: 'app,linux'
  properties: {
    serverFarmId: appServicePlan.id
    siteConfig: {
      // Node runtime
      linuxFxVersion: 'NODE|18-lts'
    }
  }
}

resource appSettings 'Microsoft.Web/sites/config@2022-03-01' = {
  parent: webApp
  name: 'appsettings'
  properties: {
    WEBSITE_RUN_FROM_PACKAGE: '1'
  }
}

output defaultHostName string = webApp.properties.defaultHostName
