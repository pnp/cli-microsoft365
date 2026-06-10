import type { ICachePlugin } from '@azure/msal-node';
import type { IPersistence } from '@azure/msal-node-extensions';
import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../utils/sinonUtil.js';
import { msalCachePlugin } from './msalCachePlugin.js';

describe('msalCachePlugin', () => {
  let mockPersistence: IPersistence;
  let mockPlugin: ICachePlugin;

  beforeEach(() => {
    msalCachePlugin.resetForTesting();

    mockPersistence = {
      save: sinon.stub().resolves(),
      load: sinon.stub().resolves(null),
      delete: sinon.stub().resolves(true),
      reloadNecessary: sinon.stub().resolves(true),
      getFilePath: sinon.stub().returns('/tmp/test-cache.json'),
      getLogger: sinon.stub().returns({
        info: () => { },
        verbose: () => { },
        error: () => { },
        warning: () => { },
        trace: () => { }
      }),
      verifyPersistence: sinon.stub().resolves(true),
      createForPersistenceValidation: sinon.stub().resolves()
    } as unknown as IPersistence;

    mockPlugin = {
      beforeCacheAccess: sinon.stub().resolves(),
      afterCacheAccess: sinon.stub().resolves()
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.unlinkSync
    ]);
    sinon.restore();
    msalCachePlugin.resetForTesting();
  });

  it(`creates persistence using PersistenceCreator`, async () => {
    const persistence = await msalCachePlugin.createPersistence();
    assert.notStrictEqual(persistence, undefined);
  });

  it(`creates plugin using PersistenceCachePlugin`, async () => {
    const plugin = await msalCachePlugin.createPlugin(mockPersistence);
    assert.notStrictEqual(plugin, undefined);
    assert.notStrictEqual(plugin.beforeCacheAccess, undefined);
    assert.notStrictEqual(plugin.afterCacheAccess, undefined);
  });

  it(`returns a cache plugin from msal-node-extensions`, async () => {
    sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin = await msalCachePlugin.getCachePlugin();
    assert.strictEqual(plugin, mockPlugin);
  });

  it(`returns the same instance on subsequent calls`, async () => {
    sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    const plugin1 = await msalCachePlugin.getCachePlugin();
    const plugin2 = await msalCachePlugin.getCachePlugin();
    assert.strictEqual(plugin1, plugin2);
  });

  it(`clears MSAL cache via persistence delete`, async () => {
    sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    await msalCachePlugin.clearMsalCache();
    assert((mockPersistence.delete as sinon.SinonStub).calledOnce);
  });

  it(`initializes persistence only once when clearing cache after getting plugin`, async () => {
    const createPersistenceStub = sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);
    sinon.stub(msalCachePlugin, 'removeLegacyCache');

    await msalCachePlugin.getCachePlugin();
    await msalCachePlugin.clearMsalCache();
    assert.strictEqual(createPersistenceStub.callCount, 1);
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
    // no assertion needed - just verifying it doesn't throw
  });

  it(`calls removeLegacyCache during initialization via getCachePlugin`, async () => {
    sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);
    const removeLegacyStub = sinon.stub(msalCachePlugin, 'removeLegacyCache');

    await msalCachePlugin.getCachePlugin();
    assert(removeLegacyStub.calledOnce);
  });

  it(`calls removeLegacyCache during initialization via clearMsalCache`, async () => {
    sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);
    const removeLegacyStub = sinon.stub(msalCachePlugin, 'removeLegacyCache');

    await msalCachePlugin.clearMsalCache();
    assert(removeLegacyStub.calledOnce);
  });
});