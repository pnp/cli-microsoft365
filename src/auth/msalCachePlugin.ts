import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import fs from 'fs';
import os from 'os';
import path from 'path';

const legacyCachePath = path.join(os.homedir(), '.cli-m365-msal.json');

const persistenceConfiguration = {
  cachePath: legacyCachePath,
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

  constructor(cachePath: string) {
    this.cachePath = cachePath;
  }

  public async beforeCacheAccess(tokenCacheContext: TokenCacheContext): Promise<void> {
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
    if (!tokenCacheContext.cacheHasChanged) {
      return;
    }

    try {
      fs.writeFileSync(this.cachePath, tokenCacheContext.tokenCache.serialize(), 'utf8');
    }
    catch {
      // Do nothing
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
        const contents = fs.readFileSync(legacyCachePath, 'utf8');
        if (contents.trim().length > 0) {
          JSON.parse(contents);
          fs.unlinkSync(legacyCachePath);
        }
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
      catch {
        // Fall back to file-based cache when native persistence is
        // unavailable (e.g. Linux without libsecret)
        return msalCachePlugin.createFileFallback();
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
      catch {
        return msalCachePlugin.createFileFallback();
      }
    })();
    const { clearCache } = await _initPromise;
    await clearCache();
  },

  resetForTesting(): void {
    _initPromise = undefined;
  }
};