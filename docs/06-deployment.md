# 6. Deployment Guide

This section covers deploying your SPFx web part and Azure Function to production environments, including CI/CD pipelines, scaling considerations, and monitoring setup.

## Prerequisites

- Azure subscription with appropriate permissions
- SharePoint Online tenant admin access
- Azure DevOps or GitHub Actions for CI/CD
- Application Insights configured

## Step 1: Azure Function Deployment

### Create Production Function App

```bash
# Create resource group
az group create --name excel-processor-prod --location eastus2

# Create storage account
az storage account create --name excelprocessorprod --location eastus2 --resource-group excel-processor-prod --sku Standard_LRS --kind StorageV2

# Create function app with premium plan for better performance
az functionapp plan create --resource-group excel-processor-prod --name excel-processor-plan --location eastus2 --number-of-workers 1 --sku EP1

# Create function app
az functionapp create --name excel-processor-prod --storage-account excelprocessorprod --plan excel-processor-plan --resource-group excel-processor-prod --runtime dotnet --functions-version 4 --os-type Windows
```

### Configure Application Settings

```bash
# Set production application settings
az functionapp config appsettings set --name excel-processor-prod --resource-group excel-processor-prod --settings \
  APP_CLIENT_ID=@Microsoft.KeyVault(SecretUri=https://excel-processor-kv.vault.azure.net/secrets/ClientId/) \
  APP_CLIENT_SECRET=@Microsoft.KeyVault(SecretUri=https://excel-processor-kv.vault.azure.net/secrets/ClientSecret/) \
  APP_TENANT_ID=@Microsoft.KeyVault(SecretUri=https://excel-processor-kv.vault.azure.net/secrets/TenantId/) \
  Environment=Production \
  ApplicationInsights__InstrumentationKey=@Microsoft.KeyVault(SecretUri=https://excel-processor-kv.vault.azure.net/secrets/AppInsightsKey/)
```

### Set Up Key Vault Access

```bash
# Get function app managed identity
FUNCTION_PRINCIPAL_ID=$(az functionapp identity show --name excel-processor-prod --resource-group excel-processor-prod --query principalId -o tsv)

# Grant access to Key Vault
az keyvault set-policy --name excel-processor-kv --object-id $FUNCTION_PRINCIPAL_ID --secret-permissions get list
```

### Deploy Function Code

#### Option A: Azure CLI Deployment

```bash
# Build and publish
dotnet publish -c Release
cd bin/Release/net6.0/publish
func azure functionapp publish excel-processor-prod --dotnet
```

#### Option B: ZIP Deploy

```bash
# Create deployment package
dotnet publish -c Release -o ./publish
cd publish
zip -r ../deployment.zip .

# Deploy
az functionapp deployment source config-zip --resource-group excel-processor-prod --name excel-processor-prod --src ../deployment.zip
```

## Step 2: SPFx Web Part Deployment

### Prepare Production Build

```bash
# Clean and build for production
gulp clean
gulp build --production
gulp bundle --production
gulp package-solution --production
```

### Upload to SharePoint App Catalog

1. Navigate to your SharePoint Admin Center
2. Go to **More features** → **Apps** → **Open**
3. Upload the `.sppkg` file from `sharepoint/solution/`
4. Check **Make this solution available to all sites in the organization**
5. Click **Deploy**

### Approve API Permissions

1. In SharePoint Admin Center → **API access**
2. Find pending requests for your app
3. Review and approve the Microsoft Graph and SharePoint permissions

### Deploy Web Part to Site

1. Navigate to your Excel Processing site
2. Edit the page with your document libraries
3. Add the Excel Processor web part
4. Configure properties:
   - Document Library ID: Your Input Files library GUID
   - Azure Function URL: `https://excel-processor-prod.azurewebsites.net/api/ProcessExcelFiles?code=[function-key]`

## Step 3: CI/CD Pipeline Setup

### Azure DevOps Pipeline

#### Create azure-pipelines.yml for Azure Function

```yaml
trigger:
  branches:
    include:
    - main
    - develop

pool:
  vmImage: 'windows-latest'

variables:
  buildConfiguration: 'Release'
  dotnetSdkVersion: '6.x'

stages:
- stage: Build
  jobs:
  - job: BuildFunction
    steps:
    - task: UseDotNet@2
      displayName: 'Use .NET SDK $(dotnetSdkVersion)'
      inputs:
        packageType: 'sdk'
        version: '$(dotnetSdkVersion)'

    - task: DotNetCoreCLI@2
      displayName: 'Restore packages'
      inputs:
        command: 'restore'
        projects: '**/*.csproj'

    - task: DotNetCoreCLI@2
      displayName: 'Build'
      inputs:
        command: 'build'
        projects: '**/*.csproj'
        arguments: '--configuration $(buildConfiguration) --no-restore'

    - task: DotNetCoreCLI@2
      displayName: 'Test'
      inputs:
        command: 'test'
        projects: '**/*Tests.csproj'
        arguments: '--configuration $(buildConfiguration) --no-build --verbosity normal'

    - task: DotNetCoreCLI@2
      displayName: 'Publish'
      inputs:
        command: 'publish'
        projects: '**/*.csproj'
        arguments: '--configuration $(buildConfiguration) --no-build --output $(Build.ArtifactStagingDirectory)'
        zipAfterPublish: true

    - publish: $(Build.ArtifactStagingDirectory)
      artifact: functionapp

- stage: Deploy
  condition: and(succeeded(), eq(variables['Build.SourceBranch'], 'refs/heads/main'))
  jobs:
  - deployment: DeployFunction
    environment: 'production'
    strategy:
      runOnce:
        deploy:
          steps:
          - task: AzureFunctionApp@1
            displayName: 'Azure Function App Deploy'
            inputs:
              azureSubscription: 'Azure-Subscription'
              appType: functionApp
              appName: 'excel-processor-prod'
              package: '$(Pipeline.Workspace)/functionapp/*.zip'
              deploymentMethod: 'zipDeploy'
```

#### Create azure-pipelines.yml for SPFx

```yaml
trigger:
  branches:
    include:
    - main
    - develop

pool:
  vmImage: 'ubuntu-latest'

variables:
  nodeVersion: '14.x'

stages:
- stage: Build
  jobs:
  - job: BuildSPFx
    steps:
    - task: NodeTool@0
      displayName: 'Use Node $(nodeVersion)'
      inputs:
        versionSpec: '$(nodeVersion)'

    - script: |
        npm ci
      displayName: 'npm ci'

    - script: |
        gulp clean
        gulp build --production
        gulp bundle --production
        gulp package-solution --production
      displayName: 'Build SPFx solution'

    - publish: sharepoint/solution
      artifact: spfxpackage

- stage: Deploy
  condition: and(succeeded(), eq(variables['Build.SourceBranch'], 'refs/heads/main'))
  jobs:
  - deployment: DeploySPFx
    environment: 'production'
    strategy:
      runOnce:
        deploy:
          steps:
          - task: AzureCLI@2
            displayName: 'Deploy to SharePoint'
            inputs:
              azureSubscription: 'Azure-Subscription'
              scriptType: 'bash'
              scriptLocation: 'inlineScript'
              inlineScript: |
                # Upload to app catalog via Microsoft Graph API
                # This requires appropriate permissions and scripting
                echo "SPFx package ready for manual upload to SharePoint App Catalog"
```

### GitHub Actions (Alternative)

#### .github/workflows/azure-function.yml

```yaml
name: Deploy Azure Function

on:
  push:
    branches: [ main ]
    paths: [ 'azure-excel-processor/**' ]
  workflow_dispatch:

env:
  AZURE_FUNCTIONAPP_NAME: 'excel-processor-prod'
  AZURE_FUNCTIONAPP_PACKAGE_PATH: 'azure-excel-processor'
  DOTNET_VERSION: '6.0.x'

jobs:
  build-and-deploy:
    runs-on: windows-latest
    steps:
    - name: 'Checkout GitHub Action'
      uses: actions/checkout@v3

    - name: Setup DotNet ${{ env.DOTNET_VERSION }} Environment
      uses: actions/setup-dotnet@v2
      with:
        dotnet-version: ${{ env.DOTNET_VERSION }}

    - name: 'Resolve Project Dependencies Using DotNet'
      shell: pwsh
      run: |
        pushd './${{ env.AZURE_FUNCTIONAPP_PACKAGE_PATH }}'
        dotnet build --configuration Release --output ./output
        popd

    - name: 'Run Azure Functions Action'
      uses: Azure/functions-action@v1
      id: fa
      with:
        app-name: ${{ env.AZURE_FUNCTIONAPP_NAME }}
        slot-name: 'production'
        package: '${{ env.AZURE_FUNCTIONAPP_PACKAGE_PATH }}/output'
        publish-profile: ${{ secrets.AZURE_FUNCTIONAPP_PUBLISH_PROFILE }}
```

## Step 4: Scaling and Performance

### Azure Function Scaling

#### Configure Auto-scaling

```bash
# Set up auto-scaling for function app
az monitor autoscale create \
  --resource /subscriptions/.../resourceGroups/excel-processor-prod/providers/Microsoft.Web/serverFarms/excel-processor-plan \
  --name excel-processor-autoscale \
  --min-count 1 \
  --max-count 10 \
  --count 1

# Add CPU percentage rule
az monitor autoscale rule create \
  --resource /subscriptions/.../resourceGroups/excel-processor-prod/providers/Microsoft.Web/serverFarms/excel-processor-plan \
  --autoscale-name excel-processor-autoscale \
  --condition "Percentage CPU > 70 avg 5m" \
  --scale out 2 \
  --cooldown 5
```

#### Premium Plan Benefits

- Faster scaling (seconds vs minutes)
- VNET integration for secure SharePoint access
- More memory and CPU options
- Persistent file system for temporary processing

### Optimize Function Performance

#### Update host.json for Production

```json
{
  "version": "2.0",
  "logging": {
    "applicationInsights": {
      "samplingSettings": {
        "isEnabled": true,
        "excludedTypes": "Request"
      },
      "enableLiveMetricsFilters": true
    }
  },
  "functionTimeout": "00:15:00",
  "extensions": {
    "http": {
      "routePrefix": "api",
      "maxOutstandingRequests": 200,
      "maxConcurrentRequests": 100
    }
  },
  "concurrency": {
    "dynamicConcurrencyEnabled": true,
    "snapshotPersistenceEnabled": true
  }
}
```

## Step 5: Monitoring and Alerting

### Application Insights Setup

```bash
# Enable Application Insights
az monitor app-insights component create \
  --app excel-processor-insights \
  --location eastus2 \
  --resource-group excel-processor-prod \
  --application-type web

# Get instrumentation key
INSTRUMENTATION_KEY=$(az monitor app-insights component show \
  --app excel-processor-insights \
  --resource-group excel-processor-prod \
  --query instrumentationKey -o tsv)

# Update function app settings
az functionapp config appsettings set \
  --name excel-processor-prod \
  --resource-group excel-processor-prod \
  --settings APPINSIGHTS_INSTRUMENTATIONKEY=$INSTRUMENTATION_KEY
```

### Configure Alerts

```bash
# Alert on function failures
az monitor metrics alert create \
  --name "Function Failures" \
  --resource /subscriptions/.../resourceGroups/excel-processor-prod/providers/Microsoft.Web/sites/excel-processor-prod \
  --description "Alert when function execution fails" \
  --condition "count > 5 where Result == Failure" \
  --window-size 5m \
  --evaluation-frequency 1m \
  --action /subscriptions/.../resourceGroups/excel-processor-prod/providers/microsoft.insights/actionGroups/myActionGroup

# Alert on high response times
az monitor metrics alert create \
  --name "Slow Response Time" \
  --resource /subscriptions/.../resourceGroups/excel-processor-prod/providers/Microsoft.Web/sites/excel-processor-prod \
  --description "Alert when response time is high" \
  --condition "avg Http5xx > 10" \
  --window-size 5m \
  --evaluation-frequency 1m
```

### Custom Metrics and Logging

Add custom telemetry to your function:

```csharp
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;

public class TelemetryHelper
{
    private static TelemetryClient _telemetry;

    static TelemetryHelper()
    {
        var configuration = TelemetryConfiguration.CreateDefault();
        configuration.InstrumentationKey = Environment.GetEnvironmentVariable("APPINSIGHTS_INSTRUMENTATIONKEY");
        _telemetry = new TelemetryClient(configuration);
    }

    public static void TrackProcessingEvent(string userId, int fileCount, TimeSpan duration, bool success)
    {
        var properties = new Dictionary<string, string>
        {
            { "UserId", userId },
            { "FileCount", fileCount.ToString() },
            { "Success", success.ToString() }
        };

        var metrics = new Dictionary<string, double>
        {
            { "ProcessingDuration", duration.TotalSeconds },
            { "FilesProcessed", fileCount }
        };

        _telemetry.TrackEvent("ExcelProcessing", properties, metrics);
        _telemetry.Flush();
    }
}
```

## Step 6: Backup and Disaster Recovery

### Function App Backup

```bash
# Configure backup for function app
az backup protection enable-for-azurewl \
  --resource-group excel-processor-prod \
  --vault-name myBackupVault \
  --item-name excel-processor-prod \
  --policy-name DefaultPolicy \
  --workload-type AzureAppService
```

### Key Vault Backup

```bash
# Backup Key Vault secrets
az keyvault secret backup \
  --vault-name excel-processor-kv \
  --name ClientId \
  --file ClientId.backup

az keyvault secret backup \
  --vault-name excel-processor-kv \
  --name ClientSecret \
  --file ClientSecret.backup
```

## Step 7: Security Hardening

### Network Security

```bash
# Configure VNET integration
az functionapp vnet-integration add \
  --resource-group excel-processor-prod \
  --name excel-processor-prod \
  --vnet MyVNet \
  --subnet MySubnet

# Add IP restrictions
az functionapp config access-restriction add \
  --resource-group excel-processor-prod \
  --name excel-processor-prod \
  --rule-name "SharePoint Access" \
  --action Allow \
  --ip-address "13.107.6.0/24" \
  --priority 100
```

### Function Authentication

```bash
# Enable function-level authentication
az functionapp config set \
  --name excel-processor-prod \
  --resource-group excel-processor-prod \
  --auth-settings '{"enabled": true, "defaultProvider": "AzureActiveDirectory"}'
```

## Step 8: Testing Production Deployment

### Smoke Tests

Create automated tests to verify production deployment:

```bash
# Test function availability
curl -f "https://excel-processor-prod.azurewebsites.net/api/ProcessExcelFiles?code=YOUR_KEY" \
  -H "Content-Type: application/json" \
  -d '{"test": true}'

# Test SharePoint connectivity (if you have a test endpoint)
curl -f "https://excel-processor-prod.azurewebsites.net/api/TestConnection"
```

### Load Testing

Use Azure Load Testing service or tools like k6:

```javascript
// k6 load test script
import http from 'k6/http';
import { check, sleep } from 'k6';

export let options = {
  stages: [
    { duration: '2m', target: 10 },  // Ramp up to 10 users
    { duration: '5m', target: 10 },  // Stay at 10 users
    { duration: '2m', target: 0 },   // Ramp down to 0
  ],
};

export default function () {
  let response = http.post(
    'https://excel-processor-prod.azurewebsites.net/api/ProcessExcelFiles?code=YOUR_KEY',
    JSON.stringify({
      siteUrl: 'https://yourtenant.sharepoint.com/sites/test',
      fileUrls: ['https://yourtenant.sharepoint.com/sites/test/Shared Documents/test.xlsx'],
      userId: 'test@example.com'
    }),
    {
      headers: {
        'Content-Type': 'application/json',
      },
    }
  );

  check(response, {
    'status is 200': (r) => r.status === 200,
    'response time < 30000': (r) => r.timings.duration < 30000,
  });

  sleep(1);
}
```

## Step 9: Rollback Strategy

### Blue-Green Deployment

```bash
# Create staging slot
az functionapp deployment slot create \
  --name excel-processor-prod \
  --resource-group excel-processor-prod \
  --slot staging

# Deploy to staging
az functionapp deployment source config-zip \
  --name excel-processor-prod \
  --resource-group excel-processor-prod \
  --slot staging \
  --src deployment.zip

# Swap slots after testing
az functionapp deployment slot swap \
  --name excel-processor-prod \
  --resource-group excel-processor-prod \
  --slot staging
```

## Step 10: Documentation and Handover

### Create Runbook

Document operational procedures:

```markdown
# Excel Processor Production Runbook

## Monitoring
- Application Insights: [link]
- Azure Monitor: [link]
- Alert notifications: [email/distribution list]

## Key Contacts
- Development Team: [contact]
- SharePoint Admins: [contact]
- Azure Subscription Owner: [contact]

## Common Issues
1. Function timeouts: Check Application Insights for long-running operations
2. Authentication errors: Verify Key Vault secrets and app registration
3. SharePoint access: Check API permissions and tenant allowlist

## Scaling Procedures
- Manual scale: Use Azure portal or CLI
- Auto-scale triggers: CPU > 70%, Queue length > 10
- Emergency scale: Contact Azure support

## Backup and Recovery
- Function code: Git repository
- Configuration: Azure Resource Manager templates
- Data: SharePoint native backup (user data)
```

## Deployment Checklist

- [ ] Azure resources created and configured
- [ ] Key Vault secrets populated
- [ ] Function app deployed and tested
- [ ] SPFx package uploaded to app catalog
- [ ] API permissions approved
- [ ] Web part configured on site
- [ ] CI/CD pipelines configured
- [ ] Monitoring and alerting set up
- [ ] Backup procedures documented
- [ ] Security policies applied
- [ ] Load testing completed
- [ ] Runbook created and distributed

## Next Steps

With deployment complete, review the [best practices guide](./07-best-practices.md) for ongoing maintenance, security, and optimization recommendations.


