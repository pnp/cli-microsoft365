import * as sinon from 'sinon';
import * as assert from 'assert';
import { fail } from 'assert';
import auth from './AzmgmtAuth';
import Auth, { Service } from '../../Auth';
import Utils from '../../Utils';
import { CommandError } from '../../Command';

describe('AzmgmtAuth', () => {
  it('restores all persisted connection properties', (done) => {
    const persistedConnection = {
      accessToken: 'abc',
      connected: true,
      resource: 'https://management.azure.com/',
      expiresAt: 123,
      refreshToken: 'def'
    };
    auth.service = new Service('https://management.azure.com/');
    sinon.stub(auth as any, 'getServiceConnectionInfo').callsFake(() => Promise.resolve(persistedConnection));
    auth
      .restoreAuth()
      .then(() => {
        try {
          assert.equal(auth.service.accessToken, persistedConnection.accessToken);
          assert.equal(auth.service.connected, persistedConnection.connected);
          assert.equal(auth.service.resource, persistedConnection.resource);
          assert.equal(auth.service.expiresAt, persistedConnection.expiresAt);
          assert.equal(auth.service.refreshToken, persistedConnection.refreshToken);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore((auth as any).getServiceConnectionInfo);
        }
      });
  });

  it('continues when restoring connection information fails', (done) => {
    sinon.stub(auth as any, 'getServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .restoreAuth()
      .then(() => {
        Utils.restore((auth as any).getServiceConnectionInfo);
        done();
      }, () => {
        Utils.restore((auth as any).getServiceConnectionInfo);
        fail('Expected promise resolve but rejected instead');
      });
  });

  it('reuses existing token if still valid', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.service = new Service('https://management.azure.com/');
    auth.service.accessToken = 'ABC';
    auth.service.expiresAt = (new Date().getTime() / 1000) + 60;
    const authEnsureAccessTokenSpy = sinon.spy(Auth.prototype, 'ensureAccessToken');
    auth
      .ensureAccessToken('https://management.azure.com', stdout)
      .then((accessToken) => {
        try {
          assert.equal(accessToken, 'ABC');
          assert(authEnsureAccessTokenSpy.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(Auth.prototype.ensureAccessToken);
        }
      });
  });

  it('retrieves new access token if previous token expired', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.service = new Service('https://management.azure.com/');
    auth.service.accessToken = 'ABC';
    auth.service.expiresAt = (new Date().getTime() / 1000) - 60;
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth
      .ensureAccessToken('https://management.azure.com', stdout)
      .then((accessToken) => {
        try {
          assert.equal(accessToken, 'DEF');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            Auth.prototype.ensureAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('retrieves new access token if no token for the specified resource available', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.service = new Service('https://management.azure.com/');
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth
      .ensureAccessToken('https://management.azure.com', stdout)
      .then((accessToken) => {
        try {
          assert.equal(accessToken, 'DEF');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            Auth.prototype.ensureAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('doesn\'t fail if persisting connection state fails', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.service = new Service('https://management.azure.com/');
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .ensureAccessToken('https://management.azure.com', stdout)
      .then((accessToken) => {
        try {
          assert.equal(accessToken, 'DEF');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            Auth.prototype.ensureAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('logs error when persisting connection state fails and running in debug mode', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.service = new Service('https://management.azure.com/');
    const stdoutLogSpy = sinon.spy(stdout, 'log');
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .ensureAccessToken('https://management.azure.com', stdout, true)
      .then((accessToken) => {
        try {
          assert(stdoutLogSpy.calledWith(new CommandError('An error has occurred')));
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            Auth.prototype.ensureAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('fails if retrieving a new access token failed', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.service = new Service('https://management.azure.com/');
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .ensureAccessToken('https://management.azure.com', stdout, true)
      .then(() => {
        Utils.restore([
          Auth.prototype.ensureAccessToken,
          (auth as any).setServiceConnectionInfo
        ]);
        fail('Failure expected but passed');
      }, (error: any) => {
        try {
          assert.equal(error, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            Auth.prototype.ensureAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('stores connection info for the Azure Management service', () => {
    const authSetServiceConnectionInfoStub = sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    const site = new Service('https://management.azure.com/');
    auth.storeConnectionInfo();
    try {
      assert(authSetServiceConnectionInfoStub.calledWith('AzMgmt', site));
    }
    catch (e) {
      throw e;
    }
    finally {
      Utils.restore((auth as any).setServiceConnectionInfo);
    }
  });

  it('clears connection info for the Azure Management service', () => {
    const authClearServiceConnectionInfoStub = sinon.stub(auth as any, 'clearServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth.clearConnectionInfo();
    try {
      assert(authClearServiceConnectionInfoStub.calledWith('AzMgmt'));
    }
    catch (e) {
      throw e;
    }
    finally {
      Utils.restore((auth as any).clearServiceConnectionInfo);
    }
  });
});