import { PublicClientApplication, SilentFlowRequest, AuthenticationResult } from '@azure/msal-node';
import { DEFAULT_CONFIG, getTokenCache, saveTokenCache, TokenCache } from './config.js';
import { authenticateWithDeviceCode } from './device-code.js';

let msalInstance: PublicClientApplication | null = null;

function getMsalInstance(): PublicClientApplication {
  if (!msalInstance) {
    if (!DEFAULT_CONFIG.clientId) {
      throw new Error('M365_CLIENT_ID environment variable is required');
    }

    msalInstance = new PublicClientApplication({
      auth: {
        clientId: DEFAULT_CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${DEFAULT_CONFIG.tenantId}`,
      },
    });
  }
  return msalInstance;
}

export async function getAccessToken(): Promise<string> {
  const cache = getTokenCache();

  // No cached token, need to authenticate
  if (!cache || !cache.account) {
    const newCache = await authenticateWithDeviceCode();
    return newCache.accessToken;
  }

  // Token still valid (with 5 min buffer)
  if (cache.expiresAt > Date.now() + 5 * 60 * 1000) {
    return cache.accessToken;
  }

  // Token expired, try silent refresh
  console.error('[Auth] Token expired, attempting silent refresh...');
  try {
    const pca = getMsalInstance();

    const silentRequest: SilentFlowRequest = {
      scopes: DEFAULT_CONFIG.scopes,
      account: {
        homeAccountId: cache.account.homeAccountId,
        environment: cache.account.environment,
        tenantId: cache.account.tenantId,
        username: cache.account.username,
        localAccountId: cache.account.homeAccountId.split('.')[0],
      },
    };

    const result: AuthenticationResult = await pca.acquireTokenSilent(silentRequest);

    console.error('[Auth] Silent refresh successful');
    const newCache: TokenCache = {
      accessToken: result.accessToken,
      refreshToken: '',
      expiresAt: result.expiresOn?.getTime() || Date.now() + 3600 * 1000,
      account: cache.account,
    };

    saveTokenCache(newCache);
    return result.accessToken;
  } catch (error) {
    // Silent refresh failed - log the actual error
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error('[Auth] Silent refresh FAILED:', errorMessage);
    console.error('[Auth] Token has expired and cannot be refreshed automatically.');
    console.error('[Auth] Please re-authenticate by running: npm run login');

    // Don't call authenticateWithDeviceCode() - it will hang in MCP context
    // Instead, throw an error so the user knows they need to re-auth
    throw new Error(
      `Token expired and refresh failed: ${errorMessage}. ` +
      `Please re-authenticate by running 'npm run login' in the ops-personal-m365-mcp directory.`
    );
  }
}

export async function isAuthenticated(): Promise<boolean> {
  const cache = getTokenCache();
  return cache !== null && cache.account !== undefined;
}

export function getCurrentUser(): string | null {
  const cache = getTokenCache();
  return cache?.account?.username || null;
}
