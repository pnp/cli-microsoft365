import * as sinon from 'sinon';
import * as assert from 'assert';
import { fail } from 'assert';
import auth, { Site } from './SpoAuth';
import { Auth } from '../../Auth';
import Utils from '../../Utils';
import { CommandError } from '../../Command';

describe('SpoAuth', () => {
  it('restores all persisted connection properties', (done) => {
    const now = new Date();
    const persistedSite = {
      tenantId: 'tid',
      url: 'https://contoso-admin.sharepoint.com',
      accessToken: 'abc',
      accessTokens: {
        'SPO': {
          accessToken: 'abc',
          expiresOn: now.toISOString()
        }
      },
      connected: true,
      resource: 'https://contoso-admin.sharepoint.com',
      expiresOn: now.toISOString(),
      refreshToken: 'def'
    };
    auth.site = new Site();
    sinon.stub(auth as any, 'getServiceConnectionInfo').callsFake(() => Promise.resolve(persistedSite));
    auth
      .restoreAuth()
      .then(() => {
        try {
          assert.equal(auth.site.tenantId, persistedSite.tenantId);
          assert.equal(auth.site.url, persistedSite.url);
          assert.equal(auth.site.accessToken, persistedSite.accessToken);
          assert.equal(auth.site.accessTokens, persistedSite.accessTokens);
          assert.equal(auth.site.connected, persistedSite.connected);
          assert.equal(auth.site.resource, persistedSite.resource);
          assert.equal(auth.site.expiresOn, persistedSite.expiresOn);
          assert.equal(auth.site.refreshToken, persistedSite.refreshToken);
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

  it('restores persisted connection properties to site and service', (done) => {
    const persistedSite = {
      tenantId: 'tid',
      url: 'https://contoso-admin.sharepoint.com',
      accessToken: 'abc',
      accessTokens: {
        'SPO': {
          accessToken: 'abc',
          expiresAt: 123
        }
      },
      connected: true,
      resource: 'https://contoso-admin.sharepoint.com',
      expiresAt: 123,
      refreshToken: 'def'
    };
    auth.site = new Site();
    sinon.stub(auth as any, 'getServiceConnectionInfo').callsFake(() => Promise.resolve(persistedSite));
    auth
      .restoreAuth()
      .then(() => {
        try {
          assert.equal(auth.site.connected, persistedSite.connected);
          assert.equal(auth.service.connected, persistedSite.connected);
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
    const expiresOn = new Date();
    expiresOn.setSeconds(expiresOn.getSeconds() + 1);
    auth.site = new Site();
    auth.site.accessTokens['https://contoso.sharepoint.com'] = {
      accessToken: 'ABC',
      expiresOn: expiresOn.toISOString()
    };
    const authEnsureAccessTokenSpy = sinon.spy(Auth.prototype, 'ensureAccessToken');
    auth
      .ensureAccessToken('https://contoso.sharepoint.com', stdout)
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
    const expiresOn = new Date();
    expiresOn.setSeconds(expiresOn.getSeconds() - 1);
    auth.site = new Site();
    auth.site.accessTokens['https://contoso.sharepoint.com'] = {
      accessToken: 'ABC',
      expiresOn: expiresOn.toISOString()
    };
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth
      .ensureAccessToken('https://contoso.sharepoint.com', stdout)
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
    auth.site = new Site();
    auth.site.accessTokens = {};
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth
      .ensureAccessToken('https://contoso.sharepoint.com', stdout)
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

  it('stores newly retrieved access token in memory for future reuse', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.site = new Site();
    auth.site.accessTokens = {};
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth
      .ensureAccessToken('https://contoso.sharepoint.com', stdout)
      .then((accessToken) => {
        try {
          assert.equal(auth.site.accessTokens['https://contoso.sharepoint.com'].accessToken, 'DEF');
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
    auth.site = new Site();
    auth.site.accessTokens = {};
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .ensureAccessToken('https://contoso.sharepoint.com', stdout)
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
    auth.site = new Site();
    auth.site.accessTokens = {};
    const stdoutLogSpy = sinon.spy(stdout, 'log');
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .ensureAccessToken('https://contoso.sharepoint.com', stdout, true)
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
    auth.site = new Site();
    auth.site.accessTokens = {};
    sinon.stub(Auth.prototype, 'ensureAccessToken').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .ensureAccessToken('https://contoso.sharepoint.com', stdout, true)
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

  it('reuses existing token if still valid when getting access token', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    const expiresOn = new Date();
    expiresOn.setSeconds(expiresOn.getSeconds() + 1);
    auth.site = new Site();
    auth.site.accessTokens['https://contoso.sharepoint.com'] = {
      accessToken: 'ABC',
      expiresOn: expiresOn.toISOString()
    };
    const authGetAccessTokenSpy = sinon.spy(Auth.prototype, 'getAccessToken');
    auth
      .getAccessToken('https://contoso.sharepoint.com', 'ABC', stdout)
      .then((accessToken) => {
        try {
          assert.equal(accessToken, 'ABC');
          assert(authGetAccessTokenSpy.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(Auth.prototype.getAccessToken);
        }
      });
  });

  it('retrieves new access token if previous token expired when getting access token', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    const expiresOn = new Date();
    expiresOn.setSeconds(expiresOn.getSeconds() - 1);
    auth.site = new Site();
    auth.site.accessTokens['https://contoso.sharepoint.com'] = {
      accessToken: 'ABC',
      expiresOn: expiresOn.toISOString()
    };
    sinon.stub(Auth.prototype, 'getAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth
      .getAccessToken('https://contoso.sharepoint.com', 'ABC', stdout)
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
            Auth.prototype.getAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('retrieves new access token if no token for the specified resource available when getting access token', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.site = new Site();
    auth.site.accessTokens = {};
    sinon.stub(Auth.prototype, 'getAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth
      .getAccessToken('https://contoso.sharepoint.com', 'ABC', stdout)
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
            Auth.prototype.getAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('stores newly retrieved access token in memory for future reuse when getting access token', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.site = new Site();
    auth.site.accessTokens = {};
    sinon.stub(Auth.prototype, 'getAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth
      .getAccessToken('https://contoso.sharepoint.com', 'ABC', stdout)
      .then((accessToken) => {
        try {
          assert.equal(auth.site.accessTokens['https://contoso.sharepoint.com'].accessToken, 'DEF');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            Auth.prototype.getAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('doesn\'t fail if persisting connection state fails when getting access token', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.site = new Site();
    auth.site.accessTokens = {};
    sinon.stub(Auth.prototype, 'getAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .getAccessToken('https://contoso.sharepoint.com', 'ABC', stdout)
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
            Auth.prototype.getAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('logs error when persisting connection state fails and running in debug mode when getting access token', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.site = new Site();
    auth.site.accessTokens = {};
    const stdoutLogSpy = sinon.spy(stdout, 'log');
    sinon.stub(Auth.prototype, 'getAccessToken').callsFake(() => Promise.resolve('DEF'));
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .getAccessToken('https://contoso.sharepoint.com', 'ABC', stdout, true)
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
            Auth.prototype.getAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('fails if getting a new access token failed', (done) => {
    const stdout = {
      log: (msg: string) => { }
    };
    auth.site = new Site();
    auth.site.accessTokens = {};
    sinon.stub(Auth.prototype, 'getAccessToken').callsFake(() => Promise.reject('An error has occurred'));
    auth
      .getAccessToken('https://contoso.sharepoint.com', 'ABC', stdout, true)
      .then(() => {
        Utils.restore([
          Auth.prototype.getAccessToken,
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
            Auth.prototype.getAccessToken,
            (auth as any).setServiceConnectionInfo
          ]);
        }
      });
  });

  it('stores connection info for the SPO service', () => {
    const authSetServiceConnectionInfoStub = sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());
    const site = new Site();
    site.url = 'https://contoso.sharepoint.com';
    auth.site = site;
    auth.storeSiteConnectionInfo();
    try {
      assert(authSetServiceConnectionInfoStub.calledWith('SPO', site));
    }
    catch (e) {
      throw e;
    }
    finally {
      Utils.restore((auth as any).setServiceConnectionInfo);
    }
  });

  it('clears connection info for the SPO service', () => {
    const authClearServiceConnectionInfoStub = sinon.stub(auth as any, 'clearServiceConnectionInfo').callsFake(() => Promise.resolve());
    auth.clearSiteConnectionInfo();
    try {
      assert(authClearServiceConnectionInfoStub.calledWith('SPO'));
    }
    catch (e) {
      throw e;
    }
    finally {
      Utils.restore((auth as any).clearServiceConnectionInfo);
    }
  });
});