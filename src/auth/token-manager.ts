import { PublicClientApplication, SilentFlowRequest, AuthenticationResult, AccountInfo } from '@azure/msal-node';
import { getAuthConfig, MSAL_CACHE_FILE, getTokenCache } from './config.js';
import { FileCachePlugin } from './cache-plugin.js';

let msalInstance: PublicClientApplication | null = null;
let msalClientId: string | null = null;
const cachePlugin = new FileCachePlugin(MSAL_CACHE_FILE);

function getMsalInstance(): PublicClientApplication {
  const config = getAuthConfig();

  // Recreate instance if client ID changed
  if (msalInstance && msalClientId !== config.clientId) {
    msalInstance = null;
  }

  if (!msalInstance) {
    if (!config.clientId) {
      throw new Error(
        'No client ID configured. Run: myoffice login --client-id <your-azure-app-client-id>\n' +
        'Or set M365_CLIENT_ID environment variable.'
      );
    }

    msalClientId = config.clientId;
    msalInstance = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
      cache: {
        cachePlugin,
      },
    });
  }
  return msalInstance;
}

export async function getAccessToken(): Promise<string> {
  const pca = getMsalInstance();

  // Get accounts from MSAL's persisted cache
  const accounts = await pca.getTokenCache().getAllAccounts();

  if (accounts.length === 0) {
    // No accounts in MSAL cache - check legacy token.json for migration hint
    const legacyCache = getTokenCache();
    if (legacyCache?.account) {
      console.error('[Auth] Found legacy token.json but no MSAL cache.');
      console.error('[Auth] Please re-authenticate to migrate to new cache format.');
    }
    throw new Error(
      'Not authenticated. Please run "npm run login" in the myoffice-mcp directory.'
    );
  }

  // Use the first account (typically there's only one for personal use)
  const account = accounts[0];

  // Try silent token acquisition (MSAL handles refresh automatically)
  console.error('[Auth] Acquiring token silently...');
  try {
    const config = getAuthConfig();
    const silentRequest: SilentFlowRequest = {
      scopes: config.scopes,
      account: account,
    };

    const result: AuthenticationResult = await pca.acquireTokenSilent(silentRequest);

    console.error('[Auth] Token acquired successfully');
    return result.accessToken;
  } catch (error) {
    // Silent refresh failed - log the actual error
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error('[Auth] Silent token acquisition FAILED:', errorMessage);
    console.error('[Auth] Please re-authenticate by running: npm run login');

    throw new Error(
      `Token acquisition failed: ${errorMessage}. ` +
      `Please re-authenticate by running 'npm run login' in the myoffice-mcp directory.`
    );
  }
}

export async function isAuthenticated(): Promise<boolean> {
  try {
    const pca = getMsalInstance();
    const accounts = await pca.getTokenCache().getAllAccounts();
    return accounts.length > 0;
  } catch {
    return false;
  }
}

export async function getCurrentUser(): Promise<string | null> {
  try {
    const pca = getMsalInstance();
    const accounts = await pca.getTokenCache().getAllAccounts();
    return accounts[0]?.username || null;
  } catch {
    return null;
  }
}
