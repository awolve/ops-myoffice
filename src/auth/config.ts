import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'fs';
import { homedir } from 'os';
import { join, dirname } from 'path';

export interface TokenCache {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  account?: {
    homeAccountId: string;
    environment: string;
    tenantId: string;
    username: string;
  };
}

export interface AuthConfig {
  clientId: string;
  tenantId: string;
  scopes: string[];
}

// Default Azure AD app for personal M365 access
// Users can override with their own app registration
export const DEFAULT_CONFIG: AuthConfig = {
  clientId: process.env.M365_CLIENT_ID || '',
  tenantId: process.env.M365_TENANT_ID || 'common',
  scopes: [
    'https://graph.microsoft.com/Mail.ReadWrite',
    'https://graph.microsoft.com/Mail.Send',
    'https://graph.microsoft.com/Calendars.ReadWrite',
    'https://graph.microsoft.com/Tasks.ReadWrite',
    'https://graph.microsoft.com/Files.ReadWrite',
    'https://graph.microsoft.com/Sites.Read.All',
    'https://graph.microsoft.com/Contacts.Read',
    'https://graph.microsoft.com/User.Read',
    'offline_access',
  ],
};

const TOKEN_FILE = join(homedir(), '.config', 'ops-personal-m365-mcp', 'token.json');

export function getTokenCache(): TokenCache | null {
  try {
    if (existsSync(TOKEN_FILE)) {
      const data = readFileSync(TOKEN_FILE, 'utf-8');
      return JSON.parse(data);
    }
  } catch {
    // Ignore errors, return null
  }
  return null;
}

export function saveTokenCache(cache: TokenCache): void {
  const dir = dirname(TOKEN_FILE);
  if (!existsSync(dir)) {
    mkdirSync(dir, { recursive: true });
  }
  writeFileSync(TOKEN_FILE, JSON.stringify(cache, null, 2), { mode: 0o600 });
}

export function clearTokenCache(): void {
  try {
    if (existsSync(TOKEN_FILE)) {
      writeFileSync(TOKEN_FILE, '', { mode: 0o600 });
    }
  } catch {
    // Ignore errors
  }
}
