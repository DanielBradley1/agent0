# Agent0

A full stack demo app showcasing **First Party Tool Calling for AI Agents** with Microsoft Entra ID authentication.

![Agent0](/docs/agent0.png "Agent0 Logo")

## What is Agent0?

Agent0 demonstrates how to build an AI-powered chat application that can securely call APIs on behalf of the logged-in user. It uses:

- **Microsoft Entra ID** for authentication (SPA + API)
- **On-Behalf-Of (OBO) flow** for the AI agent to call Microsoft Graph
- **OpenAI GPT-4** for AI capabilities
- **React + Vite** for the frontend
- **Fastify** for the backend API

Follow [this link](/docs/agent0.md) to learn more about Agent0's architecture.

## Architecture

```
┌─────────────────┐         ┌─────────────────┐         ┌─────────────────┐
│   React SPA     │  token  │  Fastify API    │  OBO    │ Microsoft Graph │
│  (Agent0 SPA)   │────────▶│  (Agent0 API)   │────────▶│    /v1.0/me     │
│                 │         │                 │         │                 │
└─────────────────┘         └─────────────────┘         └─────────────────┘
        │                           │
        │ login                     │ validates JWT
        ▼                           ▼
┌─────────────────────────────────────────────────────┐
│              Microsoft Entra ID                      │
└─────────────────────────────────────────────────────┘
```

## Pre-requisites

- [Microsoft Entra ID tenant](https://entra.microsoft.com)
- [OpenAI API Key](https://platform.openai.com/api-keys)
- Node.js version 22+
- PowerShell 7+ (for automated setup) or access to Entra admin center

## Entra ID Configuration

You need to create **two app registrations** in Microsoft Entra ID:

1. **Agent0 API** - Backend API that validates tokens and calls Microsoft Graph
2. **Agent0 SPA** - Frontend single-page application

### Option 1: PowerShell Script (Automated)

Run these commands in PowerShell to create both app registrations:

```powershell
# Install Microsoft Graph PowerShell module if not already installed
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Connect to Microsoft Graph with required permissions
Connect-MgGraph -Scopes "Application.ReadWrite.All", "DelegatedPermissionGrant.ReadWrite.All"

# Get your tenant ID
$tenantId = (Get-MgContext).TenantId
Write-Host "Tenant ID: $tenantId"

# ============================================
# Step 1: Create the Agent0 API app registration
# ============================================

$apiApp = New-MgApplication -DisplayName "Agent0 API" -SignInAudience "AzureADMyOrg"

# Set the Application ID URI
Update-MgApplication -ApplicationId $apiApp.Id -IdentifierUris @("api://$($apiApp.AppId)")

# Create the access_as_user scope
$scopeId = [Guid]::NewGuid().ToString()
$oauth2PermissionScopes = @(
    @{
        Id = $scopeId
        AdminConsentDescription = "Allows the app to access Agent0 API on behalf of the signed-in user"
        AdminConsentDisplayName = "Access Agent0 API"
        IsEnabled = $true
        Type = "User"
        Value = "access_as_user"
        UserConsentDescription = "Allow the app to access Agent0 API on your behalf"
        UserConsentDisplayName = "Access Agent0 API"
    }
)

$apiSettings = @{
    Oauth2PermissionScopes = $oauth2PermissionScopes
}

Update-MgApplication -ApplicationId $apiApp.Id -Api $apiSettings

# Add Microsoft Graph User.Read delegated permission
$graphAppId = "00000003-0000-0000-c000-000000000000"
$userReadScopeId = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"

$requiredResourceAccess = @(
    @{
        ResourceAppId = $graphAppId
        ResourceAccess = @(
            @{
                Id = $userReadScopeId
                Type = "Scope"
            }
        )
    }
)

Update-MgApplication -ApplicationId $apiApp.Id -RequiredResourceAccess $requiredResourceAccess

# Create a client secret for OBO flow
$secretCredential = @{
    DisplayName = "Agent0 API Secret"
    EndDateTime = (Get-Date).AddYears(1)
}
$apiSecret = Add-MgApplicationPassword -ApplicationId $apiApp.Id -PasswordCredential $secretCredential

# Create the service principal for the API
$apiSp = New-MgServicePrincipal -AppId $apiApp.AppId

Write-Host ""
Write-Host "=== Agent0 API Created ===" -ForegroundColor Green
Write-Host "Application (client) ID: $($apiApp.AppId)"
Write-Host "Application ID URI: api://$($apiApp.AppId)"
Write-Host "Client Secret: $($apiSecret.SecretText)" -ForegroundColor Yellow
Write-Host "⚠️  Save this secret now - it won't be shown again!" -ForegroundColor Yellow

# ============================================
# Step 2: Create the Agent0 SPA app registration
# ============================================

$spaRedirectUris = @("http://localhost:8080")

$spaApp = New-MgApplication -DisplayName "Agent0 SPA" `
    -SignInAudience "AzureADMyOrg" `
    -Spa @{ RedirectUris = $spaRedirectUris }

# Add API permissions for the SPA
$spaRequiredResourceAccess = @(
    # Microsoft Graph User.Read
    @{
        ResourceAppId = $graphAppId
        ResourceAccess = @(
            @{
                Id = $userReadScopeId
                Type = "Scope"
            }
        )
    },
    # Agent0 API access_as_user
    @{
        ResourceAppId = $apiApp.AppId
        ResourceAccess = @(
            @{
                Id = $scopeId
                Type = "Scope"
            }
        )
    }
)

Update-MgApplication -ApplicationId $spaApp.Id -RequiredResourceAccess $spaRequiredResourceAccess

# Pre-authorize the SPA to call the API (skip consent)
$preAuthorizedApps = @(
    @{
        AppId = $spaApp.AppId
        DelegatedPermissionIds = @($scopeId)
    }
)

$apiSettingsWithPreAuth = @{
    Oauth2PermissionScopes = $oauth2PermissionScopes
    PreAuthorizedApplications = $preAuthorizedApps
}

Update-MgApplication -ApplicationId $apiApp.Id -Api $apiSettingsWithPreAuth

# Create the service principal for the SPA
$spaSp = New-MgServicePrincipal -AppId $spaApp.AppId

Write-Host ""
Write-Host "=== Agent0 SPA Created ===" -ForegroundColor Green
Write-Host "Application (client) ID: $($spaApp.AppId)"

# ============================================
# Step 3: Grant admin consent
# ============================================

# Grant consent for Microsoft Graph User.Read to both apps
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"

# Grant to API
New-MgOauth2PermissionGrant -ClientId $apiSp.Id -ConsentType "AllPrincipals" `
    -ResourceId $graphSp.Id -Scope "User.Read" -ErrorAction SilentlyContinue

# Grant to SPA
New-MgOauth2PermissionGrant -ClientId $spaSp.Id -ConsentType "AllPrincipals" `
    -ResourceId $graphSp.Id -Scope "User.Read" -ErrorAction SilentlyContinue

# Grant SPA access to API
New-MgOauth2PermissionGrant -ClientId $spaSp.Id -ConsentType "AllPrincipals" `
    -ResourceId $apiSp.Id -Scope "access_as_user" -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=== Admin consent granted ===" -ForegroundColor Green

# ============================================
# Output configuration for .env file
# ============================================

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Copy these values to your .env file:" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "# Microsoft Entra ID Configuration (Backend/API)"
Write-Host "ENTRA_TENANT_ID='$tenantId'"
Write-Host "ENTRA_CLIENT_ID='$($apiApp.AppId)'"
Write-Host "ENTRA_CLIENT_SECRET='$($apiSecret.SecretText)'"
Write-Host "ENTRA_AUDIENCE='api://$($apiApp.AppId)'"
Write-Host ""
Write-Host "# Frontend Entra Configuration"
Write-Host "VITE_ENTRA_TENANT_ID='$tenantId'"
Write-Host "VITE_ENTRA_CLIENT_ID='$($spaApp.AppId)'"
Write-Host "VITE_ENTRA_API_SCOPE='api://$($apiApp.AppId)/access_as_user'"
Write-Host ""

# Disconnect
Disconnect-MgGraph
```

### Option 2: Manual Setup (Entra Admin Center)

#### Step 1: Create Agent0 API App Registration

1. Go to [entra.microsoft.com](https://entra.microsoft.com) → **Identity** → **Applications** → **App registrations**
2. Click **+ New registration**
3. Configure:
   - **Name**: `Agent0 API`
   - **Supported account types**: Single tenant
4. Click **Register**
5. Note the **Application (client) ID** and **Directory (tenant) ID**

##### Expose an API:
1. Go to **Expose an API**
2. Click **Set** next to Application ID URI → Set to `api://{client-id}`
3. Click **+ Add a scope**:
   - **Scope name**: `access_as_user`
   - **Who can consent**: Admins and users
   - **Admin consent display name**: `Access Agent0 API`
   - **Admin consent description**: `Allows the app to access Agent0 API on behalf of the signed-in user`
   - **State**: Enabled

##### Add API Permissions:
1. Go to **API permissions** → **+ Add a permission**
2. Select **Microsoft Graph** → **Delegated permissions** → `User.Read`
3. Click **Grant admin consent for [tenant]**

##### Create Client Secret:
1. Go to **Certificates & secrets** → **+ New client secret**
2. Add a description and expiry
3. **Copy the secret value immediately** (shown only once)

#### Step 2: Create Agent0 SPA App Registration

1. Go to **App registrations** → **+ New registration**
2. Configure:
   - **Name**: `Agent0 SPA`
   - **Supported account types**: Single tenant
   - **Redirect URI**: 
     - Platform: **Single-page application (SPA)**
     - URI: `http://localhost:8080`
3. Click **Register**
4. Note the **Application (client) ID**

##### Add API Permissions:
1. Go to **API permissions** → **+ Add a permission**
2. Click **My APIs** → Select **Agent0 API** → Check `access_as_user` → **Add permissions**
3. Add **Microsoft Graph** → **Delegated** → `User.Read`
4. Click **Grant admin consent for [tenant]**

#### Step 3: Pre-authorize the SPA (Optional)

This skips the consent prompt for users:

1. Go back to **Agent0 API** → **Expose an API**
2. Under **Authorized client applications**, click **+ Add a client application**
3. Enter the **Agent0 SPA** client ID
4. Check the `access_as_user` scope
5. Click **Add application**

## Installation

1. Clone the repository:
```bash
git clone https://github.com/merill/agent0-id.git
cd agent0-id
```

2. Copy the environment template and fill in your values:
```bash
cp .env.example .env
```

3. Edit `.env` with your Entra ID configuration (from the setup steps above):
```env
# Microsoft Entra ID Configuration (Backend/API)
ENTRA_TENANT_ID='your-tenant-id'
ENTRA_CLIENT_ID='your-api-client-id'
ENTRA_CLIENT_SECRET='your-api-client-secret'
ENTRA_AUDIENCE='api://your-api-client-id'

# Frontend Entra Configuration
VITE_ENTRA_TENANT_ID='your-tenant-id'
VITE_ENTRA_CLIENT_ID='your-spa-client-id'
VITE_ENTRA_API_SCOPE='api://your-api-client-id/access_as_user'

# OpenAI Configuration
OPENAI_API_KEY='your-openai-api-key'
```

4. Install dependencies:
```bash
# Install server dependencies
cd server
npm install
cd ..

# Install client dependencies
cd client
npm install
cd ..
```

## Running Agent0

### Start the Server
Open a terminal and run:
```bash
cd server
npm run dev
```

### Start the Client
Open another terminal and run:
```bash
cd client
npm run dev
```

## Using Agent0

1. Open [http://localhost:8080](http://localhost:8080)
2. Click **Log In** to authenticate with Microsoft Entra ID
3. Start a new chat

### Test First-Party Tool Calling

Try asking the AI agent about yourself:

> "Who am I? Give me all the details you have"

The agent will use the `getUserInfo` tool to call Microsoft Graph on your behalf and return your profile information.

## Features

- ✅ Sign in / Sign out with Microsoft Entra ID
- ✅ Protected API with JWT validation
- ✅ On-Behalf-Of (OBO) flow for Microsoft Graph calls
- ✅ First-party tool calling (getUserInfo)
- ✅ Streaming AI responses
- ✅ Chat history persistence
- ✅ Light/Dark theme

## Troubleshooting

### 401 Unauthorized on API calls
- Ensure the SPA has API permission to call the Agent0 API
- Check that admin consent has been granted
- Clear browser localStorage and re-login

### OBO token exchange fails
- Verify `ENTRA_CLIENT_SECRET` is set correctly
- Ensure Agent0 API has `User.Read` Graph permission with admin consent
- Check the server logs for detailed error messages

### CORS errors
- Ensure `ALLOWED_ORIGINS` in `.env` includes `http://localhost:8080`

## Learn More

- [Microsoft Entra Agent ID Platform](https://learn.microsoft.com/en-us/entra/agent-id/identity-platform/what-is-agent-id-platform)
- [On-Behalf-Of Flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow)
- [MSAL React Documentation](https://learn.microsoft.com/en-us/entra/identity-platform/tutorial-v2-react)
