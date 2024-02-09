import { ISerializableTokenCache, TokenCacheContext } from '@azure/msal-node';
import assert from 'assert';
import sinon from 'sinon';
import { sinonUtil } from '../utils/sinonUtil.js';
import { msalCachePlugin } from './msalCachePlugin.js';

const mockCache: ISerializableTokenCache = {
  deserialize: () => { },
  serialize: () => ''
};
const mockCacheContext = new TokenCacheContext(mockCache, false);

describe('msalCachePlugin', () => {
  let mockCacheDeserializeSpy: sinon.SinonSpy;
  let mockCacheSerializeSpy: sinon.SinonSpy;

  before(() => {
    mockCacheDeserializeSpy = sinon.spy(mockCache, 'deserialize');
    mockCacheSerializeSpy = sinon.spy(mockCache, 'serialize');
  });

  afterEach(() => {
    mockCacheDeserializeSpy.resetHistory();
    mockCacheSerializeSpy.resetHistory();
    mockCacheContext.hasChanged = false;
    sinonUtil.restore([
      (msalCachePlugin as any).fileTokenStorage.get,
      (msalCachePlugin as any).fileTokenStorage.set
    ]);
  });

  it(`restores token cache from the cache storage`, async () => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'get').returns('');
    await msalCachePlugin.beforeCacheAccess(mockCacheContext);
    assert(mockCacheDeserializeSpy.called);
  });

  it(`doesn't fail restoring cache if cache file not found`, async () => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'get').rejects('File not found');
    await msalCachePlugin.beforeCacheAccess(mockCacheContext);
    assert(mockCacheDeserializeSpy.notCalled);
  });

  it(`doesn't fail restoring cache if an error has occurred`, async () => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'get').rejects('An error has occurred');
    await msalCachePlugin.beforeCacheAccess(mockCacheContext);
    assert(mockCacheDeserializeSpy.notCalled);
  });

  it(`persists cache on disk when cache changed`, async () => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'set').resolves();
    mockCacheContext.hasChanged = true;

    await msalCachePlugin.afterCacheAccess(mockCacheContext);
    assert(mockCacheSerializeSpy.called);
  });

  it(`doesn't persist cache on disk when cache not changed`, async () => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'set').resolves();
    await msalCachePlugin.afterCacheAccess(mockCacheContext);
    assert(mockCacheSerializeSpy.notCalled);
  });

  it(`doesn't throw exception when persisting cache failed`, async () => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'set').rejects('An error has occurred');
    mockCacheContext.hasChanged = true;
    await msalCachePlugin.afterCacheAccess(mockCacheContext);
    assert(mockCacheSerializeSpy.called);
  });
});