import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import fs from 'fs';
import os from 'os';
import path from 'path';

const legacyCachePath = path.join(os.homedir(), '.cli-m365-msal.json');
const fallbackCachePath = path.join(os.homedir(), '.cli-m365-msal-cache.json');

const persistenceConfiguration = {
  cachePath: fallbackCachePath,
  serviceName: 'cli-microsoft365',
  accountName: 'msal-cache',
  usePlaintextFileOnLinux: true
};

let _initPromise: Promise<{ plugin: ICachePlugin; clearCache: () => Promise<void> }> | undefined;

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
  private lockFileHandle: number | undefined;

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
      this.releaseLock();
    }
  }

  private async acquireLock(): Promise<void> {
    const lockPath = `${this.cachePath}.lockfile`;

    for (let retry = 0; retry < 500; retry++) {
      try {
        this.lockFileHandle = fs.openSync(lockPath, 'wx', 0o600);
        return;
      }
      catch (err) {
        const errorCode = (err as NodeJS.ErrnoException).code;
        if (errorCode !== 'EEXIST' && errorCode !== 'EPERM') {
          throw err;
        }
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }

    throw new Error(`Could not acquire MSAL cache lock at ${lockPath}`);
  }

  private releaseLock(): void {
    if (this.lockFileHandle === undefined) {
      return;
    }

    const lockFileHandle = this.lockFileHandle;
    this.lockFileHandle = undefined;
    try {
      fs.unlinkSync(`${this.cachePath}.lockfile`);
    }
    finally {
      fs.closeSync(lockFileHandle);
    }
  }
}

export const msalCachePlugin = {
  async importMsalExtensions(): Promise<typeof import('@azure/msal-node-extensions')> {
    return await import('@azure/msal-node-extensions');
  },

  async createNativePersistence(): Promise<{ plugin: ICachePlugin; clearCache: () => Promise<void> }> {
    const { DataProtectionScope, PersistenceCachePlugin, PersistenceCreator } = await msalCachePlugin.importMsalExtensions();
    const persistence = await PersistenceCreator.createPersistence({
      ...persistenceConfiguration,
      dataProtectionScope: DataProtectionScope.CurrentUser
    });
    return {
      plugin: new PersistenceCachePlugin(persistence),
      clearCache: async () => { await persistence.delete(); }
    };
  },

  createFileFallback(): { plugin: ICachePlugin; clearCache: () => Promise<void> } {
    return {
      plugin: new FileCachePlugin(persistenceConfiguration.cachePath),
      clearCache: async () => {
        try { fs.unlinkSync(persistenceConfiguration.cachePath); }
        catch { /* file may not exist */ }
      }
    };
  },

  removeLegacyCache(): void {
    try {
      if (fs.existsSync(legacyCachePath)) {
        fs.unlinkSync(legacyCachePath);
      }
    }
    catch {
      // Ignore errors: file may already be managed by the new
      // persistence layer (e.g. DPAPI-encrypted on Windows)
    }
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
    const { clearCache } = await _initPromise;
    await clearCache();
    // Also remove the file-based fallback cache to ensure no tokens
    // remain if the machine previously used file-based persistence
    try { fs.unlinkSync(persistenceConfiguration.cachePath); }
    catch { /* file may not exist */ }
  },

  resetForTesting(): void {
    _initPromise = undefined;
  }
};