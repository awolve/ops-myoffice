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

export interface StoredConfig {
  clientId?: string;
  tenantId?: string;
}

const CONFIG_DIR = join(homedir(), '.config', 'myoffice-mcp');
const CONFIG_FILE = join(CONFIG_DIR, 'config.json');

/**
 * Get stored configuration from config file
 */
export function getStoredConfig(): StoredConfig {
  try {
    if (existsSync(CONFIG_FILE)) {
      const data = readFileSync(CONFIG_FILE, 'utf-8');
      return JSON.parse(data);
    }
  } catch {
    // Ignore errors, return empty config
  }
  return {};
}

/**
 * Save configuration to config file
 */
export function saveStoredConfig(config: StoredConfig): void {
  if (!existsSync(CONFIG_DIR)) {
    mkdirSync(CONFIG_DIR, { recursive: true });
  }
  // Merge with existing config
  const existing = getStoredConfig();
  const merged = { ...existing, ...config };
  writeFileSync(CONFIG_FILE, JSON.stringify(merged, null, 2), { mode: 0o600 });
}

/**
 * Get the client ID from env var or stored config
 */
function getClientId(): string {
  // Environment variable takes precedence
  if (process.env.M365_CLIENT_ID) {
    return process.env.M365_CLIENT_ID;
  }
  // Fall back to stored config
  const stored = getStoredConfig();
  return stored.clientId || '';
}

/**
 * Get the tenant ID from env var or stored config
 */
function getTenantId(): string {
  // Environment variable takes precedence
  if (process.env.M365_TENANT_ID) {
    return process.env.M365_TENANT_ID;
  }
  // Fall back to stored config, default to 'common'
  const stored = getStoredConfig();
  return stored.tenantId || 'common';
}

// Scopes for Graph API access
const SCOPES = [
  'https://graph.microsoft.com/Mail.ReadWrite',
  'https://graph.microsoft.com/Mail.Send',
  'https://graph.microsoft.com/Calendars.ReadWrite',
  'https://graph.microsoft.com/Tasks.ReadWrite',
  'https://graph.microsoft.com/Files.ReadWrite',
  'https://graph.microsoft.com/Sites.ReadWrite.All',
  'https://graph.microsoft.com/Contacts.ReadWrite',
  'https://graph.microsoft.com/User.Read',
  // Teams
  'https://graph.microsoft.com/Team.ReadBasic.All',
  'https://graph.microsoft.com/Channel.ReadBasic.All',
  'https://graph.microsoft.com/ChannelMessage.Read.All',
  'https://graph.microsoft.com/ChannelMessage.Send',
  // Chats
  'https://graph.microsoft.com/Chat.Create',
  'https://graph.microsoft.com/Chat.ReadBasic',
  'https://graph.microsoft.com/Chat.Read',
  'https://graph.microsoft.com/ChatMessage.Send',
  // Planner
  'https://graph.microsoft.com/Tasks.ReadWrite',
  'https://graph.microsoft.com/Group.Read.All',
  'https://graph.microsoft.com/User.ReadBasic.All',
  'offline_access',
];

/**
 * Get the current auth config (dynamically evaluated)
 * Call this function instead of using DEFAULT_CONFIG directly
 * to ensure env vars and stored config are always checked
 */
export function getAuthConfig(): AuthConfig {
  return {
    clientId: getClientId(),
    tenantId: getTenantId(),
    scopes: SCOPES,
  };
}

// Default Azure AD app for personal M365 access
// Users can override with their own app registration
// NOTE: This is evaluated once at module load time.
// For dynamic access, use getAuthConfig() instead.
export const DEFAULT_CONFIG: AuthConfig = {
  clientId: getClientId(),
  tenantId: getTenantId(),
  scopes: SCOPES,
};

const TOKEN_FILE = join(homedir(), '.config', 'myoffice-mcp', 'token.json');
export const MSAL_CACHE_FILE = join(homedir(), '.config', 'myoffice-mcp', 'msal-cache.json');

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
