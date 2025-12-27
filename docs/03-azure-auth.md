# 3. Azure Authentication Setup

This section covers setting up Azure Active Directory (AAD) authentication for secure communication between your SharePoint SPFx web part and Azure Function.

## Prerequisites

- Azure subscription with admin access
- SharePoint Online tenant
- Application Administrator or Global Administrator role in Azure AD

## Architecture Overview

```
┌─────────────────┐    Bearer Token    ┌──────────────────┐
│                 │ ─────────────────► │                  │
│  SPFx Web Part  │                    │  Azure Function  │
│  (Client)       │ ◄────────────────  │  (API)           │
│                 │   API Response     │                  │
└─────────────────┘                    └──────────────────┘
         │                                    │
         │   Client Credentials               │ SharePoint API
         ▼   Flow (App-Only)                  ▼
    ┌─────────────────┐                 ┌──────────────────┐
    │                 │                 │                  │
    │  Azure AD App   │                 │  SharePoint      │
    │  Registration   │                 │  Online          │
    │                 │                 │                  │
    └─────────────────┘                 └──────────────────┘
```

## Step 1: Register Azure AD Application

### Create App Registration

1. Navigate to **Azure Portal** → **Azure Active Directory** → **App registrations**
2. Click **New registration**
3. Configure:
   ```
   Name: Excel Processor API
   Supported account types: Accounts in this organizational directory only
   Redirect URI: Web → https://yourtenant.sharepoint.com (leave blank for now)
   ```
4. Click **Register**

### Record Application Details

After registration, note these values:
- **Application (client) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
- **Directory (tenant) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

## Step 2: Configure API Permissions

### Add Microsoft Graph Permissions

1. In your app registration → **API permissions** → **Add a permission**
2. Select **Microsoft Graph**
3. Add these **Application permissions**:
   - `Sites.ReadWrite.All` - Read and write items in all site collections
   - `Files.ReadWrite.All` - Read and write files in all site collections

### Add SharePoint Permissions

4. Click **Add a permission** → **APIs my organization uses**
5. Search for and select **SharePoint**
6. Add **Application permissions**:
   - `Sites.ReadWrite.All` - Read and write items in all site collections

### Grant Admin Consent

7. Click **Grant admin consent for [your tenant]**
8. Confirm the consent dialog

## Step 3: Create Client Secret

### Generate Secret

1. Go to **Certificates & secrets** → **New client secret**
2. Configure:
   ```
   Description: Excel Processor Secret
   Expires: 24 months (recommended for production)
   ```
3. Click **Add**
4. **Important**: Copy the secret value immediately (it won't be shown again)

### Store Secret Securely

- Store in Azure Key Vault (production)
- Use Azure Function Application Settings (development)
- Never commit to source control

## Step 4: Configure Authentication for Azure Function

### Option A: App-Only Authentication (Recommended)

The Azure Function will use client credentials flow to authenticate with SharePoint.

### Option B: On-Behalf-Of Flow (Alternative)

For delegated permissions (if needed for user-specific operations).

## Step 5: Create Azure Key Vault (Production)

### Set Up Key Vault

1. **Azure Portal** → **Key Vaults** → **Create**
2. Configure:
   ```
   Name: excel-processor-kv
   Region: Same as your resources
   Pricing tier: Standard
   ```
3. Create access policy for your Azure Function's managed identity

### Store Secrets

1. **Secrets** → **Generate/Import**
2. Add secrets:
   - `ClientId` → Your Application ID
   - `ClientSecret` → Your client secret
   - `TenantId` → Your tenant ID

## Step 6: Configure Azure Function Authentication

### Function App Settings

In your Azure Function → **Configuration** → **Application settings**, add:

```
APP_CLIENT_ID=your-application-id
APP_CLIENT_SECRET=your-client-secret
APP_TENANT_ID=your-tenant-id
```

### Use Key Vault References (Production)

For production, use Key Vault references:

```
APP_CLIENT_ID=@Microsoft.KeyVault(SecretUri=https://excel-processor-kv.vault.azure.net/secrets/ClientId/)
APP_CLIENT_SECRET=@Microsoft.KeyVault(SecretUri=https://excel-processor-kv.vault.azure.net/secrets/ClientSecret/)
APP_TENANT_ID=@Microsoft.KeyVault(SecretUri=https://excel-processor-kv.vault.azure.net/secrets/TenantId/)
```

## Step 7: Implement Authentication in Azure Function

### C# Authentication Code

Create an authentication helper class:

```csharp
using Microsoft.Identity.Client;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace ExcelProcessor.Auth
{
    public class SharePointAuthenticator
    {
        private readonly IConfidentialClientApplication _app;

        public SharePointAuthenticator(string clientId, string clientSecret, string tenantId)
        {
            _app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();
        }

        public async Task<string> GetAccessTokenAsync()
        {
            string[] scopes = { "https://yourtenant.sharepoint.com/.default" };

            try
            {
                var result = await _app.AcquireTokenForClient(scopes).ExecuteAsync();
                return result.AccessToken;
            }
            catch (MsalServiceException ex)
            {
                throw new Exception($"Error acquiring access token: {ex.Message}", ex);
            }
        }

        public async Task<HttpClient> GetAuthenticatedHttpClientAsync()
        {
            var token = await GetAccessTokenAsync();
            var client = new HttpClient();

            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            return client;
        }
    }
}
```

### Usage in Function

```csharp
[FunctionName("ProcessExcelFiles")]
public async Task<IActionResult> Run(
    [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
    ILogger log)
{
    try
    {
        // Get configuration
        var clientId = Environment.GetEnvironmentVariable("APP_CLIENT_ID");
        var clientSecret = Environment.GetEnvironmentVariable("APP_CLIENT_SECRET");
        var tenantId = Environment.GetEnvironmentVariable("APP_TENANT_ID");

        // Create authenticator
        var authenticator = new SharePointAuthenticator(clientId, clientSecret, tenantId);

        // Get authenticated HTTP client
        var httpClient = await authenticator.GetAuthenticatedHttpClientAsync();

        // Use httpClient for SharePoint API calls
        // ... rest of your function logic

    }
    catch (Exception ex)
    {
        log.LogError(ex, "Error in ProcessExcelFiles function");
        return new BadRequestObjectResult(new { error = ex.Message });
    }
}
```

## Step 8: Test Authentication

### Create Test Function

Create a simple test endpoint to verify authentication:

```csharp
[FunctionName("TestSharePointConnection")]
public async Task<IActionResult> TestConnection(
    [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
    ILogger log)
{
    try
    {
        var authenticator = new SharePointAuthenticator(
            Environment.GetEnvironmentVariable("APP_CLIENT_ID"),
            Environment.GetEnvironmentVariable("APP_CLIENT_SECRET"),
            Environment.GetEnvironmentVariable("APP_TENANT_ID")
        );

        var token = await authenticator.GetAccessTokenAsync();

        // Test SharePoint API call
        var httpClient = await authenticator.GetAuthenticatedHttpClientAsync();
        var siteUrl = "https://yourtenant.sharepoint.com/sites/excel-processing";

        var response = await httpClient.GetAsync($"{siteUrl}/_api/web");
        response.EnsureSuccessStatusCode();

        var content = await response.Content.ReadAsStringAsync();

        return new OkObjectResult(new
        {
            message = "Authentication successful",
            hasToken = !string.IsNullOrEmpty(token),
            sharePointAccessible = true
        });

    }
    catch (Exception ex)
    {
        log.LogError(ex, "Authentication test failed");
        return new BadRequestObjectResult(new
        {
            error = ex.Message,
            hasToken = false,
            sharePointAccessible = false
        });
    }
}
```

### Test the Function

1. Deploy the test function
2. Call the endpoint: `GET https://yourfunction.azurewebsites.net/api/TestSharePointConnection`
3. Verify response shows successful authentication

## Step 9: Secure Function Access

### Function Authorization

Configure your Azure Function with proper authorization:

1. **Function** → **Authorization** → Set to **Function** level
2. Generate function keys for your SPFx web part

### Get Function URL

Your function URL will be:
```
https://yourfunction.azurewebsites.net/api/ProcessExcelFiles?code=your-function-key
```

### Configure SPFx Web Part

Update your SPFx web part properties with the secured function URL.

## Step 10: Implement Token Caching (Optional)

For better performance, implement token caching:

```csharp
using Microsoft.Extensions.Caching.Memory;

public class TokenCacheProvider : ITokenCacheProvider
{
    private readonly IMemoryCache _cache;

    public TokenCacheProvider(IMemoryCache cache)
    {
        _cache = cache;
    }

    public void SetCache(string key, byte[] value)
    {
        _cache.Set(key, value, TimeSpan.FromMinutes(50)); // Token lifetime - 5min buffer
    }

    public byte[] GetCache(string key)
    {
        return _cache.Get<byte[]>(key);
    }
}
```

## Security Best Practices

### ✅ Do's

- Use Azure Key Vault for secrets in production
- Implement proper error handling without exposing sensitive information
- Use managed identities when possible
- Rotate secrets regularly
- Monitor authentication failures
- Use least privilege permissions

### ❌ Don'ts

- Hardcode secrets in source code
- Use overly broad permissions
- Log sensitive information
- Share client secrets between environments
- Use personal accounts for app registrations

## Troubleshooting

### Common Authentication Issues

1. **"Invalid client" error**: Check client ID and secret
2. **"Insufficient privileges"**: Verify API permissions and admin consent
3. **Token expired**: Implement proper token refresh logic
4. **SharePoint access denied**: Check site permissions for the app

### Debug Authentication

```csharp
// Add detailed logging
log.LogInformation($"Client ID: {clientId?.Substring(0, 8)}...");
log.LogInformation($"Tenant ID: {tenantId}");

// Test token acquisition
try
{
    var token = await authenticator.GetAccessTokenAsync();
    log.LogInformation("Token acquired successfully");
}
catch (Exception ex)
{
    log.LogError(ex, "Token acquisition failed");
}
```

## Next Steps

With authentication configured, proceed to [building the Azure Function](./04-azure-function.md) for Excel processing and report generation.

## Configuration Summary

| Setting | Development | Production |
|---------|-------------|------------|
| Client Secret | App Settings | Key Vault |
| Function Auth | Function Key | Function Key + IP Restrictions |
| Logging | Console | Application Insights |
| Monitoring | Local Debug | Azure Monitor |



