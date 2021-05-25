import { ISerializableTokenCache, TokenCacheContext } from '@azure/msal-node';
import * as assert from 'assert';
import * as sinon from 'sinon';
import Utils from '../Utils';
import { msalCachePlugin } from './msalCachePlugin';

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
    Utils.restore([
      (msalCachePlugin as any).fileTokenStorage.get,
      (msalCachePlugin as any).fileTokenStorage.set
    ]);
  });

  it(`restores token cache from the cache storage`, (done) => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'get').callsFake(() => '');
    msalCachePlugin
      .beforeCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheDeserializeSpy.called);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't fail restoring cache if cache file not found`, (done) => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'get').callsFake(() => Promise.reject('File not found'));
    msalCachePlugin
      .beforeCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheDeserializeSpy.notCalled);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't fail restoring cache if cache file not found`, (done) => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'get').callsFake(() => Promise.reject('File not found'));
    msalCachePlugin
      .beforeCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheDeserializeSpy.notCalled);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't fail restoring cache if an error has occurred`, (done) => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'get').callsFake(() => Promise.reject('An error has occurred'));
    msalCachePlugin
      .beforeCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheDeserializeSpy.notCalled);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`persists cache on disk when cache changed`, (done) => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'set').callsFake(() => Promise.resolve());
    mockCacheContext.hasChanged = true;
    msalCachePlugin
      .afterCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheSerializeSpy.called);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't persist cache on disk when cache not changed`, (done) => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'set').callsFake(() => Promise.resolve());
    msalCachePlugin
      .afterCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheSerializeSpy.notCalled);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't throw exception when persisting cache failed`, (done) => {
    sinon.stub((msalCachePlugin as any).fileTokenStorage, 'set').callsFake(() => Promise.reject('An error has occurred'));
    mockCacheContext.hasChanged = true;
    msalCachePlugin
      .afterCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheSerializeSpy.called);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });
});