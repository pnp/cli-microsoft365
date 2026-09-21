import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { lock } from 'proper-lockfile';

const legacyCachePath = path.join(os.homedir(), '.cli-m365-msal.json');
const fallbackCachePath = path.join(os.homedir(), '.cli-m365-msal-cache.json');
const lockRetryCount = 500;
const lockRetryDelay = 100;
const staleLockThreshold = 10000;

const persistenceConfiguration = {
  cachePath: fallbackCachePath,
  serviceName: 'cli-microsoft365',
  accountName: 'msal-cache',
  usePlaintextFileOnLinux: true
};

export const fileLock = { lock };

let _initPromise: Promise<{ plugin: ICachePlugin; clearCache: () => Promise<void>; isFileFallback: boolean }> | undefined;

// Fallback ICachePlugin that stores tokens as plain JSON on disk.
// @azure/msal-node-extensions ships a usePlaintextFileOnLinux option
// for Linux systems without libsecret, but the package's barrel export
// eagerly loads LibSecretPersistence which does `import keytar` at the
// top level. When libsecret is missing, the entire dynamic import of
// the package fails before any fallback logic can run. This class
// provides equivalent file-based persistence without depending on the
// package at all.
// See: https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/7170
class FileCachePlugin implements ICachePlugin {
  private cachePath: string;
  private releaseCacheLock: (() => Promise<void>) | undefined;

  constructor(cachePath: string) {
    this.cachePath = cachePath;
  }

  public async beforeCacheAccess(tokenCacheContext: TokenCacheContext): Promise<void> {
    await this.acquireLock();

    try {
      if (fs.existsSync(this.cachePath)) {
        const data = fs.readFileSync(this.cachePath, 'utf8');
        tokenCacheContext.tokenCache.deserialize(data);
      }
    }
    catch {
      // Do nothing
    }
  }

  public async afterCacheAccess(tokenCacheContext: TokenCacheContext): Promise<void> {
    try {
      if (tokenCacheContext.cacheHasChanged) {
        if (fs.existsSync(this.cachePath)) {
          fs.chmodSync(this.cachePath, 0o600);
        }
        fs.writeFileSync(this.cachePath, tokenCacheContext.tokenCache.serialize(), { encoding: 'utf8', mode: 0o600 });
      }
    }
    catch {
      // Do nothing
    }
    finally {
      await this.releaseLock();
    }
  }

  public async clearCache(): Promise<void> {
    await this.acquireLock();

    try {
      removeFile(this.cachePath);
    }
    finally {
      await this.releaseLock();
    }
  }

  private async acquireLock(): Promise<void> {
    this.releaseCacheLock = await msalCachePlugin.acquireFileLock(this.cachePath);
  }

  private async releaseLock(): Promise<void> {
    if (this.releaseCacheLock === undefined) {
      return;
    }

    const releaseCacheLock = this.releaseCacheLock;
    this.releaseCacheLock = undefined;
    await releaseCacheLock();
  }
}

function removeFile(filePath: string): void {
  try {
    fs.unlinkSync(filePath);
  }
  catch (err) {
    if ((err as NodeJS.ErrnoException).code !== 'ENOENT') {
      throw err;
    }
  }
}

export const msalCachePlugin = {
  async acquireFileLock(cachePath: string): Promise<() => Promise<void>> {
    return await fileLock.lock(cachePath, {
      lockfilePath: `${cachePath}.lockfile`,
      realpath: false,
      stale: staleLockThreshold,
      update: staleLockThreshold / 2,
      retries: {
        retries: lockRetryCount,
        factor: 1,
        minTimeout: lockRetryDelay,
        maxTimeout: lockRetryDelay,
        randomize: false
      }
    });
  },

  async importMsalExtensions(): Promise<typeof import('@azure/msal-node-extensions')> {
    return await import('@azure/msal-node-extensions');
  },

  async createNativePersistence(): Promise<{ plugin: ICachePlugin; clearCache: () => Promise<void>; isFileFallback: boolean }> {
    const { DataProtectionScope, PersistenceCachePlugin, PersistenceCreator } = await msalCachePlugin.importMsalExtensions();
    const persistence = await PersistenceCreator.createPersistence({
      ...persistenceConfiguration,
      dataProtectionScope: DataProtectionScope.CurrentUser
    });
    return {
      plugin: new PersistenceCachePlugin(persistence),
      clearCache: async () => { await persistence.delete(); },
      isFileFallback: false
    };
  },

  createFileFallback(): { plugin: ICachePlugin; clearCache: () => Promise<void>; isFileFallback: boolean } {
    const plugin = new FileCachePlugin(persistenceConfiguration.cachePath);
    return {
      plugin,
      clearCache: async () => { await plugin.clearCache(); },
      isFileFallback: true
    };
  },

  removeLegacyCache(): void {
    removeFile(legacyCachePath);
  },

  async getCachePlugin(): Promise<ICachePlugin> {
    _initPromise ??= (async () => {
      msalCachePlugin.removeLegacyCache();
      try {
        return await msalCachePlugin.createNativePersistence();
      }
      catch (err) {
        // Fall back to file-based cache only on Linux where libsecret
        // may not be installed. On Windows (DPAPI) and macOS (Keychain)
        // native persistence should always work.
        if (process.platform === 'linux') {
          return msalCachePlugin.createFileFallback();
        }
        throw err;
      }
    })();
    const { plugin } = await _initPromise;
    return plugin;
  },

  async clearMsalCache(): Promise<void> {
    _initPromise ??= (async () => {
      msalCachePlugin.removeLegacyCache();
      try {
        return await msalCachePlugin.createNativePersistence();
      }
      catch (err) {
        if (process.platform === 'linux') {
          return msalCachePlugin.createFileFallback();
        }
        throw err;
      }
    })();
    const { clearCache, isFileFallback } = await _initPromise;
    await clearCache();
    // Also remove the file-based fallback cache to ensure no tokens
    // remain if the machine previously used file-based persistence
    if (!isFileFallback) {
      await msalCachePlugin.createFileFallback().clearCache();
    }
  },

  resetForTesting(): void {
    _initPromise = undefined;
  }
};