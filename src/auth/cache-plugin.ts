import { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'fs';
import { dirname } from 'path';

/**
 * MSAL Cache Plugin that persists the full token cache (including refresh tokens)
 * to a file. This ensures tokens survive server restarts.
 */
export class FileCachePlugin implements ICachePlugin {
  private cachePath: string;

  constructor(cachePath: string) {
    this.cachePath = cachePath;
  }

  /**
   * Called before MSAL accesses the cache.
   * Load the cache from disk into MSAL's memory.
   */
  async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    try {
      if (existsSync(this.cachePath)) {
        const cacheData = readFileSync(this.cachePath, 'utf-8');
        if (cacheData && cacheData.trim()) {
          cacheContext.tokenCache.deserialize(cacheData);
        }
      }
    } catch (error) {
      console.error('[Cache] Failed to load cache:', error instanceof Error ? error.message : error);
    }
  }

  /**
   * Called after MSAL accesses the cache.
   * If the cache changed, persist it to disk.
   */
  async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    if (cacheContext.cacheHasChanged) {
      try {
        const dir = dirname(this.cachePath);
        if (!existsSync(dir)) {
          mkdirSync(dir, { recursive: true });
        }
        const serialized = cacheContext.tokenCache.serialize();
        writeFileSync(this.cachePath, serialized, { mode: 0o600 });
      } catch (error) {
        console.error('[Cache] Failed to save cache:', error instanceof Error ? error.message : error);
      }
    }
  }
}
