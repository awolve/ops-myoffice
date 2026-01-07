# MSAL Cache Persistence

> Retroactive spec - documented after implementation

## Summary

Fixed token refresh by properly persisting MSAL's internal cache (including refresh tokens) to disk. Previously, refresh tokens were lost on server restart, causing silent token refresh to fail and the MCP to hang.

## Problem

The original implementation:
1. Created a new `PublicClientApplication` on each server start with empty cache
2. Saved only account info to `token.json`, not MSAL's full cache
3. When token expired, `acquireTokenSilent` failed (no refresh token in cache)
4. Fallback to device code auth hung indefinitely in MCP context

## Solution

Implement MSAL's `ICachePlugin` interface to persist the full token cache:
- `beforeCacheAccess`: Load cache from disk into MSAL memory
- `afterCacheAccess`: Save cache to disk if changed

## Changes Made

- `src/auth/cache-plugin.ts` (new): `FileCachePlugin` implementing `ICachePlugin`
- `src/auth/config.ts`: Added `MSAL_CACHE_FILE` export
- `src/auth/token-manager.ts`:
  - Uses cache plugin in `PublicClientApplication` config
  - Gets accounts from MSAL cache instead of legacy `token.json`
  - `getCurrentUser()` now async
- `src/auth/device-code.ts`: Uses cache plugin for login
- `src/index.ts`: Updated for async `getCurrentUser()`

## Behavior

1. On login (`npm run login`): MSAL cache saved to `~/.config/ops-personal-m365-mcp/msal-cache.json`
2. On server start: Cache loaded from disk, accounts available
3. On token refresh: MSAL uses cached refresh token, updates cache automatically
4. Cache file has restrictive permissions (0600)

## Tasks

- [x] Create FileCachePlugin implementing ICachePlugin
- [x] Update PublicClientApplication to use cache plugin
- [x] Update token-manager to get accounts from MSAL cache
- [x] Update device-code to use cache plugin
- [x] Update index.ts for async getCurrentUser()
- [ ] Test token refresh after server restart

## Notes

- Legacy `token.json` kept for migration hint but no longer used
- Version bumped to 0.3.0
- Users must re-authenticate once to populate new cache format
