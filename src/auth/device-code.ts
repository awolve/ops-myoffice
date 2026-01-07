import { PublicClientApplication, DeviceCodeRequest } from '@azure/msal-node';
import { DEFAULT_CONFIG, MSAL_CACHE_FILE } from './config.js';
import { FileCachePlugin } from './cache-plugin.js';

const cachePlugin = new FileCachePlugin(MSAL_CACHE_FILE);

export async function authenticateWithDeviceCode(): Promise<void> {
  if (!DEFAULT_CONFIG.clientId) {
    throw new Error(
      'M365_CLIENT_ID environment variable is required.\n' +
      'Create an Azure AD app registration with delegated permissions and set M365_CLIENT_ID.'
    );
  }

  const pca = new PublicClientApplication({
    auth: {
      clientId: DEFAULT_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${DEFAULT_CONFIG.tenantId}`,
    },
    cache: {
      cachePlugin,
    },
  });

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

  // The cache plugin automatically persists the tokens via afterCacheAccess
  console.log(`\nAuthenticated as: ${result.account?.username}`);
  console.log('Token cached at:', MSAL_CACHE_FILE);
  console.log('Refresh token is now properly persisted.\n');
}
