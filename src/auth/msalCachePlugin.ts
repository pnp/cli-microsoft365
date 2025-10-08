import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { FileTokenStorage } from './FileTokenStorage.js';
import { TokenStorage } from './TokenStorage.js';

class MsalCachePlugin implements ICachePlugin {
  private fileTokenStorage: TokenStorage = new FileTokenStorage(FileTokenStorage.msalCacheFilePath());

  public async beforeCacheAccess(tokenCacheContext: TokenCacheContext): Promise<void> {
    try {
      const data: string = await this.fileTokenStorage.get();
      tokenCacheContext.tokenCache.deserialize(data);
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
      await this.fileTokenStorage.set(tokenCacheContext.tokenCache.serialize());
    }
    catch {
      // Do nothing
    }
  }
}

const msalCachePlugin = new MsalCachePlugin();
export { msalCachePlugin };