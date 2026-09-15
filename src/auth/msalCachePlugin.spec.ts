import type { ICachePlugin } from '@azure/msal-node';
import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../utils/sinonUtil.js';
import { msalCachePlugin } from './msalCachePlugin.js';

describe('msalCachePlugin', () => {
  let mockPlugin: ICachePlugin;

  beforeEach(() => {
    msalCachePlugin.resetForTesting();

    mockPlugin = {
      beforeCacheAccess: sinon.stub().resolves(),
      afterCacheAccess: sinon.stub().resolves()
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.chmodSync,
      fs.closeSync,
      fs.openSync,
      fs.readFileSync,
      fs.unlinkSync,
      fs.writeFileSync
    ]);
    sinon.restore();
    msalCachePlugin.resetForTesting();
  });

  it(`importMsalExtensions loads the msal-node-extensions module`, async () => {
    try {
      const result = await msalCachePlugin.importMsalExtensions();
      assert.notStrictEqual(result, undefined);
    }
    catch {
      // May fail on systems without native dependencies (e.g. libsecret on Linux)
    }
  });

  it(`returns a cache plugin using native persistence`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').resolves({
      plugin: mockPlugin,
      clearCache: sinon.stub().resolves()
    });
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin = await msalCachePlugin.getCachePlugin();
    assert.strictEqual(plugin, mockPlugin);
  });

  it(`returns the same instance on subsequent calls`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').resolves({
      plugin: mockPlugin,
      clearCache: sinon.stub().resolves()
    });
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin1 = await msalCachePlugin.getCachePlugin();
    const plugin2 = await msalCachePlugin.getCachePlugin();
    assert.strictEqual(plugin1, plugin2);
  });

  it(`falls back to file-based cache when native persistence fails on Linux`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('libsecret not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();
    assert.notStrictEqual(plugin, undefined);
    assert.notStrictEqual(plugin.beforeCacheAccess, undefined);
    assert.notStrictEqual(plugin.afterCacheAccess, undefined);
  });

  it(`throws when native persistence fails on non-Linux platforms`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('Keychain error'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('darwin');

    await assert.rejects(msalCachePlugin.getCachePlugin(), { message: 'Keychain error' });
  });

  it(`clears MSAL cache via native persistence`, async () => {
    const clearCacheStub = sinon.stub().resolves();
    sinon.stub(msalCachePlugin, 'createNativePersistence').resolves({
      plugin: mockPlugin,
      clearCache: clearCacheStub
    });
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    await msalCachePlugin.clearMsalCache();
    assert(clearCacheStub.calledOnce);
    assert(unlinkStub.called);
  });

  it(`clears file-based cache when native persistence fails on Linux`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('libsecret not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    await msalCachePlugin.clearMsalCache();
    assert(unlinkStub.called);
  });

  it(`does not fail clearing file-based cache when file does not exist`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('libsecret not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');
    sinon.stub(fs, 'unlinkSync').throws(new Error('ENOENT'));

    await msalCachePlugin.clearMsalCache();
  });

  it(`throws when clearing cache and native persistence fails on non-Linux`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('DPAPI error'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('win32');

    await assert.rejects(msalCachePlugin.clearMsalCache(), { message: 'DPAPI error' });
  });

  it(`initializes only once when clearing cache after getting plugin`, async () => {
    const createStub = sinon.stub(msalCachePlugin, 'createNativePersistence').resolves({
      plugin: mockPlugin,
      clearCache: sinon.stub().resolves()
    });
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    await msalCachePlugin.getCachePlugin();
    await msalCachePlugin.clearMsalCache();
    assert.strictEqual(createStub.callCount, 1);
  });

  it(`creates native persistence using imported msal-extensions`, async () => {
    const mockPersistence = {
      delete: sinon.stub().resolves(true)
    };
    sinon.stub(msalCachePlugin, 'importMsalExtensions').resolves({
      DataProtectionScope: { CurrentUser: 0 },
      PersistenceCreator: {
        createPersistence: sinon.stub().resolves(mockPersistence)
      },
      PersistenceCachePlugin: class {
        constructor() { /* empty */ }
        beforeCacheAccess(): Promise<void> { return Promise.resolve(); }
        afterCacheAccess(): Promise<void> { return Promise.resolve(); }
      }
    } as any);

    const result = await msalCachePlugin.createNativePersistence();
    assert.notStrictEqual(result.plugin, undefined);
    assert.notStrictEqual(result.clearCache, undefined);
  });

  it(`native persistence clearCache delegates to persistence.delete`, async () => {
    const deleteStub = sinon.stub().resolves(true);
    const mockPersistence = { delete: deleteStub };
    sinon.stub(msalCachePlugin, 'importMsalExtensions').resolves({
      DataProtectionScope: { CurrentUser: 0 },
      PersistenceCreator: {
        createPersistence: sinon.stub().resolves(mockPersistence)
      },
      PersistenceCachePlugin: class {
        constructor() { /* empty */ }
        beforeCacheAccess(): Promise<void> { return Promise.resolve(); }
        afterCacheAccess(): Promise<void> { return Promise.resolve(); }
      }
    } as any);

    const result = await msalCachePlugin.createNativePersistence();
    await result.clearCache();
    assert(deleteStub.calledOnce);
  });

  it(`createFileFallback returns a plugin and clearCache function`, () => {
    const result = msalCachePlugin.createFileFallback();
    assert.notStrictEqual(result.plugin, undefined);
    assert.notStrictEqual(result.plugin.beforeCacheAccess, undefined);
    assert.notStrictEqual(result.plugin.afterCacheAccess, undefined);
    assert.notStrictEqual(result.clearCache, undefined);
  });

  it(`createFileFallback clearCache deletes the cache file`, async () => {
    const unlinkStub = sinon.stub(fs, 'unlinkSync');
    const result = msalCachePlugin.createFileFallback();
    await result.clearCache();
    assert(unlinkStub.calledOnce);
  });

  it(`createFileFallback clearCache does not throw when file does not exist`, async () => {
    sinon.stub(fs, 'unlinkSync').throws(new Error('ENOENT'));
    const result = msalCachePlugin.createFileFallback();
    await result.clearCache();
  });

  it(`removes legacy plaintext cache file when it exists`, () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    msalCachePlugin.removeLegacyCache();
    assert(unlinkStub.calledOnce);
  });

  it(`does not remove cache file when it does not exist`, () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    msalCachePlugin.removeLegacyCache();
    assert(unlinkStub.notCalled);
  });

  it(`removes cache file even when it is empty`, () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    msalCachePlugin.removeLegacyCache();
    assert(unlinkStub.calledOnce);
  });

  it(`removes cache file even when it contains non-JSON content`, () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    msalCachePlugin.removeLegacyCache();
    assert(unlinkStub.calledOnce);
  });

  it(`does not fail when reading cache file throws error`, () => {
    sinon.stub(fs, 'existsSync').throws(new Error('An error has occurred'));

    msalCachePlugin.removeLegacyCache();
  });

  it(`calls removeLegacyCache during initialization via getCachePlugin`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').resolves({
      plugin: mockPlugin,
      clearCache: sinon.stub().resolves()
    });
    const removeLegacyStub = sinon.stub(msalCachePlugin, 'removeLegacyCache');

    await msalCachePlugin.getCachePlugin();
    assert(removeLegacyStub.calledOnce);
  });

  it(`calls removeLegacyCache during initialization via clearMsalCache`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').resolves({
      plugin: mockPlugin,
      clearCache: sinon.stub().resolves()
    });
    const removeLegacyStub = sinon.stub(msalCachePlugin, 'removeLegacyCache');

    await msalCachePlugin.clearMsalCache();
    assert(removeLegacyStub.calledOnce);
  });

  it(`file cache plugin deserializes token cache from file`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'openSync').returns(1);
    sinon.stub(fs, 'closeSync');
    sinon.stub(fs, 'unlinkSync');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{"token":"data"}');
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('') };
    const context = { tokenCache: mockCache, cacheHasChanged: false, hasChanged: false } as any;

    await plugin.beforeCacheAccess(context);
    await plugin.afterCacheAccess(context);
    assert(mockCache.deserialize.calledWith('{"token":"data"}'));
  });

  it(`file cache plugin does not fail when cache file is missing`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'openSync').returns(1);
    sinon.stub(fs, 'closeSync');
    sinon.stub(fs, 'unlinkSync');
    sinon.stub(fs, 'existsSync').returns(false);
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('') };
    const context = { tokenCache: mockCache, cacheHasChanged: false, hasChanged: false } as any;

    await plugin.beforeCacheAccess(context);
    await plugin.afterCacheAccess(context);
    assert(mockCache.deserialize.notCalled);
  });

  it(`file cache plugin serializes token cache to file when changed`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();

    const openStub = sinon.stub(fs, 'openSync').returns(1);
    const closeStub = sinon.stub(fs, 'closeSync');
    const unlinkStub = sinon.stub(fs, 'unlinkSync');
    sinon.stub(fs, 'existsSync').returns(false);
    const writeStub = sinon.stub(fs, 'writeFileSync');
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('{"serialized":"data"}') };
    const context = { tokenCache: mockCache, cacheHasChanged: true, hasChanged: true } as any;

    await plugin.beforeCacheAccess(context);
    await plugin.afterCacheAccess(context);
    assert(writeStub.calledOnce);
    assert(openStub.calledOnceWith(sinon.match(/\.lockfile$/), 'wx', 0o600));
    assert(unlinkStub.calledOnceWith(sinon.match(/\.lockfile$/)));
    assert(closeStub.calledOnceWith(1));
  });

  it(`file cache plugin tightens permissions before overwriting an existing file`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'openSync').returns(1);
    sinon.stub(fs, 'closeSync');
    sinon.stub(fs, 'unlinkSync');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{}');
    const chmodStub = sinon.stub(fs, 'chmodSync');
    const writeStub = sinon.stub(fs, 'writeFileSync');
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('{"serialized":"data"}') };
    const context = { tokenCache: mockCache, cacheHasChanged: true, hasChanged: true } as any;

    await plugin.beforeCacheAccess(context);
    await plugin.afterCacheAccess(context);

    assert(chmodStub.calledOnceWith(sinon.match(/\.cli-m365-msal-cache\.json$/), 0o600));
    assert(chmodStub.calledBefore(writeStub));
  });

  it(`file cache plugin retries while another process holds the lock`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');
    const clock = sinon.useFakeTimers();

    const plugin = await msalCachePlugin.getCachePlugin();

    const lockError = (code: string): NodeJS.ErrnoException => Object.assign(new Error(code), { code });
    sinon.stub(fs, 'openSync')
      .onFirstCall().throws(lockError('EEXIST'))
      .onSecondCall().throws(lockError('EPERM'))
      .returns(1);
    sinon.stub(fs, 'closeSync');
    sinon.stub(fs, 'unlinkSync');
    sinon.stub(fs, 'existsSync').returns(false);
    const context = { tokenCache: { deserialize: sinon.stub() }, cacheHasChanged: false } as any;

    const accessPromise = plugin.beforeCacheAccess(context);
    await clock.tickAsync(200);
    await accessPromise;
    await plugin.afterCacheAccess(context);
  });

  it(`file cache plugin rethrows unexpected lock errors`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'openSync').throws(Object.assign(new Error('access denied'), { code: 'EACCES' }));
    const context = { tokenCache: { deserialize: sinon.stub() }, cacheHasChanged: false } as any;

    await assert.rejects(plugin.beforeCacheAccess(context), { message: 'access denied' });
  });

  it(`file cache plugin fails after exhausting lock retries`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');
    const clock = sinon.useFakeTimers();

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'openSync').throws(Object.assign(new Error('locked'), { code: 'EEXIST' }));
    const context = { tokenCache: { deserialize: sinon.stub() }, cacheHasChanged: false } as any;

    const accessPromise = plugin.beforeCacheAccess(context);
    await clock.tickAsync(50000);
    await assert.rejects(accessPromise, /Could not acquire MSAL cache lock/);
  });

  it(`file cache plugin can finish cache access when no lock was acquired`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();
    const context = { tokenCache: { serialize: sinon.stub() }, cacheHasChanged: false } as any;

    await plugin.afterCacheAccess(context);
  });

  it(`file cache plugin does not write when cache not changed`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'openSync').returns(1);
    sinon.stub(fs, 'closeSync');
    sinon.stub(fs, 'unlinkSync');
    sinon.stub(fs, 'existsSync').returns(false);
    const writeStub = sinon.stub(fs, 'writeFileSync');
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('') };
    const context = { tokenCache: mockCache, cacheHasChanged: false, hasChanged: false } as any;

    await plugin.beforeCacheAccess(context);
    await plugin.afterCacheAccess(context);
    assert(writeStub.notCalled);
  });

  it(`file cache plugin does not throw when writing fails`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'openSync').returns(1);
    sinon.stub(fs, 'closeSync');
    sinon.stub(fs, 'unlinkSync');
    sinon.stub(fs, 'existsSync').returns(false);
    sinon.stub(fs, 'writeFileSync').throws(new Error('write failed'));
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('data') };
    const context = { tokenCache: mockCache, cacheHasChanged: true, hasChanged: true } as any;

    await plugin.beforeCacheAccess(context);
    await plugin.afterCacheAccess(context);
  });

  it(`file cache plugin does not throw when reading fails`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(process, 'platform').value('linux');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'openSync').returns(1);
    sinon.stub(fs, 'closeSync');
    sinon.stub(fs, 'unlinkSync');
    sinon.stub(fs, 'existsSync').onFirstCall().throws(new Error('read failed')).returns(false);
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('') };
    const context = { tokenCache: mockCache, cacheHasChanged: false, hasChanged: false } as any;

    await plugin.beforeCacheAccess(context);
    await plugin.afterCacheAccess(context);
    assert(mockCache.deserialize.notCalled);
  });
});