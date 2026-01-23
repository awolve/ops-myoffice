import { PublicClientApplication, DeviceCodeRequest } from '@azure/msal-node';
import { getAuthConfig, MSAL_CACHE_FILE } from './config.js';
import { FileCachePlugin } from './cache-plugin.js';

const cachePlugin = new FileCachePlugin(MSAL_CACHE_FILE);

export async function authenticateWithDeviceCode(): Promise<void> {
  const config = getAuthConfig();

  if (!config.clientId) {
    throw new Error(
      'No client ID configured.\n' +
      'Run: myoffice login --client-id <your-azure-app-client-id>\n' +
      'Or set M365_CLIENT_ID environment variable.'
    );
  }

  const pca = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
    cache: {
      cachePlugin,
    },
  });

  const deviceCodeRequest: DeviceCodeRequest = {
    scopes: config.scopes,
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
