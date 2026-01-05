import { PublicClientApplication, DeviceCodeRequest, AuthenticationResult } from '@azure/msal-node';
import { DEFAULT_CONFIG, saveTokenCache, TokenCache } from './config.js';

export async function authenticateWithDeviceCode(): Promise<TokenCache> {
  if (!DEFAULT_CONFIG.clientId) {
    throw new Error(
      'M365_CLIENT_ID environment variable is required.\n' +
      'Create an Azure AD app registration with delegated permissions and set M365_CLIENT_ID.'
    );
  }

  const msalConfig = {
    auth: {
      clientId: DEFAULT_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${DEFAULT_CONFIG.tenantId}`,
    },
  };

  const pca = new PublicClientApplication(msalConfig);

  const deviceCodeRequest: DeviceCodeRequest = {
    scopes: DEFAULT_CONFIG.scopes,
    deviceCodeCallback: (response) => {
      console.log('\n' + '='.repeat(60));
      console.log('AUTHENTICATION REQUIRED');
      console.log('='.repeat(60));
      console.log(`\n${response.message}\n`);
      console.log('='.repeat(60) + '\n');
    },
  };

  const result = await pca.acquireTokenByDeviceCode(deviceCodeRequest);

  if (!result || !result.accessToken) {
    throw new Error('Failed to acquire access token');
  }

  const tokenCache: TokenCache = {
    accessToken: result.accessToken,
    refreshToken: '', // MSAL handles refresh internally
    expiresAt: result.expiresOn?.getTime() || Date.now() + 3600 * 1000,
    account: result.account ? {
      homeAccountId: result.account.homeAccountId,
      environment: result.account.environment,
      tenantId: result.account.tenantId,
      username: result.account.username,
    } : undefined,
  };

  saveTokenCache(tokenCache);

  console.log(`\nAuthenticated as: ${result.account?.username}`);
  console.log('Token cached for future sessions.\n');

  return tokenCache;
}
