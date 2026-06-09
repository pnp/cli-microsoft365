import type { ICachePlugin } from '@azure/msal-node';
import type { IPersistence } from '@azure/msal-node-extensions';
import assert from 'assert';
import sinon from 'sinon';
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

    const plugin = await msalCachePlugin.getCachePlugin();
    assert.strictEqual(plugin, mockPlugin);
  });

  it(`returns the same instance on subsequent calls`, async () => {
    sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);

    const plugin1 = await msalCachePlugin.getCachePlugin();
    const plugin2 = await msalCachePlugin.getCachePlugin();
    assert.strictEqual(plugin1, plugin2);
  });

  it(`clears MSAL cache via persistence delete`, async () => {
    sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);

    await msalCachePlugin.clearMsalCache();
    assert((mockPersistence.delete as sinon.SinonStub).calledOnce);
  });

  it(`initializes persistence only once when clearing cache after getting plugin`, async () => {
    const createPersistenceStub = sinon.stub(msalCachePlugin, 'createPersistence').resolves(mockPersistence);
    sinon.stub(msalCachePlugin, 'createPlugin').resolves(mockPlugin);

    await msalCachePlugin.getCachePlugin();
    await msalCachePlugin.clearMsalCache();
    assert.strictEqual(createPersistenceStub.callCount, 1);
  });
});