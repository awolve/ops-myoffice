import { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { existsSync, readFileSync, writeFileSync, mkdirSync, renameSync, unlinkSync } from 'fs';
import { dirname } from 'path';

/**
 * MSAL Cache Plugin that persists the full token cache (including refresh tokens)
 * to a file. This ensures tokens survive server restarts.
 */
export class FileCachePlugin implements ICachePlugin {
  private cachePath: string;
  private writeLock: Promise<void> = Promise.resolve();

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
   * If the cache changed, persist it to disk using atomic write.
   */
  async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    if (cacheContext.cacheHasChanged) {
      // Queue writes to prevent concurrent file operations
      this.writeLock = this.writeLock.then(() => this.writeCache(cacheContext));
      await this.writeLock;
    }
  }

  private async writeCache(cacheContext: TokenCacheContext): Promise<void> {
    try {
      const dir = dirname(this.cachePath);
      if (!existsSync(dir)) {
        mkdirSync(dir, { recursive: true });
      }
      const serialized = cacheContext.tokenCache.serialize();

      // Atomic write: write to temp file, then rename
      // This prevents corruption from concurrent writes
      const tempPath = `${this.cachePath}.tmp.${process.pid}`;
      writeFileSync(tempPath, serialized, { mode: 0o600 });
      renameSync(tempPath, this.cachePath);
    } catch (error) {
      console.error('[Cache] Failed to save cache:', error instanceof Error ? error.message : error);
      // Clean up temp file if it exists
      try {
        const tempPath = `${this.cachePath}.tmp.${process.pid}`;
        if (existsSync(tempPath)) {
          unlinkSync(tempPath);
        }
      } catch {
        // Ignore cleanup errors
      }
    }
  }
}
