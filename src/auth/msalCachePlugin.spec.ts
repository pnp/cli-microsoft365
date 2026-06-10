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
      fs.readFileSync,
      fs.unlinkSync,
      fs.writeFileSync
    ]);
    sinon.restore();
    msalCachePlugin.resetForTesting();
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

  it(`falls back to file-based cache when native persistence fails`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('libsecret not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin = await msalCachePlugin.getCachePlugin();
    assert.notStrictEqual(plugin, undefined);
    assert.notStrictEqual(plugin.beforeCacheAccess, undefined);
    assert.notStrictEqual(plugin.afterCacheAccess, undefined);
  });

  it(`clears MSAL cache via native persistence`, async () => {
    const clearCacheStub = sinon.stub().resolves();
    sinon.stub(msalCachePlugin, 'createNativePersistence').resolves({
      plugin: mockPlugin,
      clearCache: clearCacheStub
    });
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    await msalCachePlugin.clearMsalCache();
    assert(clearCacheStub.calledOnce);
  });

  it(`clears file-based cache when native persistence fails`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('libsecret not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    await msalCachePlugin.clearMsalCache();
    assert(unlinkStub.called);
  });

  it(`does not fail clearing file-based cache when file does not exist`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('libsecret not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');
    sinon.stub(fs, 'unlinkSync').throws(new Error('ENOENT'));

    await msalCachePlugin.clearMsalCache();
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

  it(`removes legacy plaintext cache file when it exists`, () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{"AccessToken":{}}');
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

  it(`does not remove cache file when it is empty`, () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('');
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    msalCachePlugin.removeLegacyCache();
    assert(unlinkStub.notCalled);
  });

  it(`does not fail when cache file contains non-JSON content`, () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('not-json-content');
    const unlinkStub = sinon.stub(fs, 'unlinkSync');

    msalCachePlugin.removeLegacyCache();
    assert(unlinkStub.notCalled);
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

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{"token":"data"}');
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('') };
    const context = { tokenCache: mockCache, cacheHasChanged: false, hasChanged: false } as any;

    await plugin.beforeCacheAccess(context);
    assert(mockCache.deserialize.calledWith('{"token":"data"}'));
  });

  it(`file cache plugin does not fail when cache file is missing`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'existsSync').returns(false);
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('') };
    const context = { tokenCache: mockCache, cacheHasChanged: false, hasChanged: false } as any;

    await plugin.beforeCacheAccess(context);
    assert(mockCache.deserialize.notCalled);
  });

  it(`file cache plugin serializes token cache to file when changed`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin = await msalCachePlugin.getCachePlugin();

    const writeStub = sinon.stub(fs, 'writeFileSync');
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('{"serialized":"data"}') };
    const context = { tokenCache: mockCache, cacheHasChanged: true, hasChanged: true } as any;

    await plugin.afterCacheAccess(context);
    assert(writeStub.calledOnce);
  });

  it(`file cache plugin does not write when cache not changed`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin = await msalCachePlugin.getCachePlugin();

    const writeStub = sinon.stub(fs, 'writeFileSync');
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('') };
    const context = { tokenCache: mockCache, cacheHasChanged: false, hasChanged: false } as any;

    await plugin.afterCacheAccess(context);
    assert(writeStub.notCalled);
  });

  it(`file cache plugin does not throw when writing fails`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'writeFileSync').throws(new Error('write failed'));
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('data') };
    const context = { tokenCache: mockCache, cacheHasChanged: true, hasChanged: true } as any;

    await plugin.afterCacheAccess(context);
  });

  it(`file cache plugin does not throw when reading fails`, async () => {
    sinon.stub(msalCachePlugin, 'createNativePersistence').rejects(new Error('not available'));
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin = await msalCachePlugin.getCachePlugin();

    sinon.stub(fs, 'existsSync').throws(new Error('read failed'));
    const mockCache = { deserialize: sinon.stub(), serialize: sinon.stub().returns('') };
    const context = { tokenCache: mockCache, cacheHasChanged: false, hasChanged: false } as any;

    await plugin.beforeCacheAccess(context);
    assert(mockCache.deserialize.notCalled);
  });
});