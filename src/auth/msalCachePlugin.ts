import type { ICachePlugin } from '@azure/msal-node';
import type { IPersistence } from '@azure/msal-node-extensions';
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

let _initPromise: Promise<{ plugin: ICachePlugin; persistence: IPersistence }> | undefined;

export const msalCachePlugin = {
  async createPersistence(): Promise<IPersistence> {
    const { DataProtectionScope, FilePersistence, PersistenceCreator } = await import('@azure/msal-node-extensions');
    try {
      return await PersistenceCreator.createPersistence({
        ...persistenceConfiguration,
        dataProtectionScope: DataProtectionScope.CurrentUser
      });
    }
    catch {
      // PersistenceCreator fails on Linux when libsecret is not installed
      // because the keytar native module cannot load. Fall back to an
      // unencrypted file, which matches the usePlaintextFileOnLinux intent.
      return FilePersistence.create(persistenceConfiguration.cachePath);
    }
  },

  async createPlugin(persistence: IPersistence): Promise<ICachePlugin> {
    const { PersistenceCachePlugin } = await import('@azure/msal-node-extensions');
    return new PersistenceCachePlugin(persistence);
  },

  removeLegacyCache(): void {
    try {
      if (fs.existsSync(legacyCachePath)) {
        const contents = fs.readFileSync(legacyCachePath, 'utf8');
        // Legacy cache is a plain JSON object with token data.
        // If parsing succeeds and the file has content, it's a
        // legacy plaintext cache that should be removed.
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
      const persistence = await msalCachePlugin.createPersistence();
      const plugin = await msalCachePlugin.createPlugin(persistence);
      return { plugin, persistence };
    })();
    const { plugin } = await _initPromise;
    return plugin;
  },

  async clearMsalCache(): Promise<void> {
    _initPromise ??= (async () => {
      msalCachePlugin.removeLegacyCache();
      const persistence = await msalCachePlugin.createPersistence();
      const plugin = await msalCachePlugin.createPlugin(persistence);
      return { plugin, persistence };
    })();
    const { persistence } = await _initPromise;
    await persistence.delete();
  },

  resetForTesting(): void {
    _initPromise = undefined;
  }
};