/// <reference types="vite/client" />

import {
  Configuration,
  PublicClientApplication,
  AccountInfo,
  InteractionRequiredAuthError,
} from '@azure/msal-browser'

// MSAL configuration
const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_ENTRA_CLIENT_ID || '',
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_ENTRA_TENANT_ID || 'common'}`,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
}

// Create the MSAL instance
export const msalInstance = new PublicClientApplication(msalConfig)

// Scopes for login and API access
export const loginRequest = {
  scopes: ['openid', 'profile', 'email'],
}

export const apiRequest = {
  scopes: [import.meta.env.VITE_ENTRA_API_SCOPE || 'api://agent0-api/.default'],
}

// Helper to get access token silently
export async function getAccessToken(
  account: AccountInfo | null
): Promise<string | null> {
  if (!account) {
    return null
  }

  try {
    const response = await msalInstance.acquireTokenSilent({
      ...apiRequest,
      account,
    })
    return response.accessToken
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      // Fallback to redirect if silent fails
      try {
        await msalInstance.acquireTokenRedirect(apiRequest)
        return null // Will redirect, so return null
      } catch (redirectError) {
        console.error('Failed to acquire token via redirect:', redirectError)
        return null
      }
    }
    console.error('Failed to acquire token silently:', error)
    return null
  }
}

// User type for compatibility with existing components
export interface EntraUser {
  sub?: string
  name?: string
  email?: string
  picture?: string
}

// Convert MSAL AccountInfo to EntraUser
export function accountToUser(account: AccountInfo | null): EntraUser | undefined {
  if (!account) return undefined

  return {
    sub: account.localAccountId,
    name: account.name || account.username,
    email: account.username,
    picture: undefined, // Entra doesn't provide picture in token, would need Graph API call
  }
}
