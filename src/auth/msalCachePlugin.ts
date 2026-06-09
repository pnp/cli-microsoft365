import type { ICachePlugin } from '@azure/msal-node';
import type { IPersistence } from '@azure/msal-node-extensions';
import os from 'os';
import path from 'path';

const persistenceConfiguration = {
  cachePath: path.join(os.homedir(), '.cli-m365-msal.json'),
  serviceName: 'cli-microsoft365',
  accountName: 'msal-cache',
  usePlaintextFileOnLinux: true
};

let _initPromise: Promise<{ plugin: ICachePlugin; persistence: IPersistence }> | undefined;

export const msalCachePlugin = {
  async createPersistence(): Promise<IPersistence> {
    const { DataProtectionScope, PersistenceCreator } = await import('@azure/msal-node-extensions');
    return PersistenceCreator.createPersistence({
      ...persistenceConfiguration,
      dataProtectionScope: DataProtectionScope.CurrentUser
    });
  },

  async createPlugin(persistence: IPersistence): Promise<ICachePlugin> {
    const { PersistenceCachePlugin } = await import('@azure/msal-node-extensions');
    return new PersistenceCachePlugin(persistence);
  },

  async getCachePlugin(): Promise<ICachePlugin> {
    _initPromise ??= (async () => {
      const persistence = await msalCachePlugin.createPersistence();
      const plugin = await msalCachePlugin.createPlugin(persistence);
      return { plugin, persistence };
    })();
    const { plugin } = await _initPromise;
    return plugin;
  },

  async clearMsalCache(): Promise<void> {
    _initPromise ??= (async () => {
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