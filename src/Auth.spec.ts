import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
import { Auth, Service, AuthType } from './Auth';
import * as os from 'os';
import { KeychainTokenStorage } from './auth/KeychainTokenStorage';
import { WindowsTokenStorage } from './auth/WindowsTokenStorage';
import { FileTokenStorage } from './auth/FileTokenStorage';
import { TokenStorage } from './auth/TokenStorage';
import { CommandError } from './Command';

class MockTokenStorage implements TokenStorage {
  public get(): Promise<string> {
    return Promise.resolve('ABC');
  }

  public set(connectionInfo: string): Promise<void> {
    return Promise.resolve();
  }

  public remove(): Promise<void> {
    return Promise.resolve();
  }
}

describe('Auth', () => {
  let log: any[];
  let auth: Auth;
  const resource: string = 'https://contoso.sharepoint.com';
  const appId: string = '9bc3ab49-b65d-410a-85ad-de819febfddc';
  const refreshToken: string = 'ref';
  const stdout: any = {
    log: (msg: any) => {
      log.push(msg);
    }
  }
  const stdoutLogSpy = sinon.spy(stdout, 'log');

  beforeEach(() => {
    log = [];
    auth = new Auth();
    (auth as any).appId = appId;
  });

  afterEach(() => {
  });

  it('returns existing access token if still valid', (done) => {
    const now = new Date();
    now.setSeconds(now.getSeconds() + 1);
    auth.service.accessTokens[resource] = {
      expiresOn: now.toISOString(),
      value: 'abc'
    }
    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      try {
        assert.equal(accessToken, auth.service.accessTokens[resource].value);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('returns existing access token if still valid (debug)', (done) => {
    const now = new Date();
    now.setSeconds(now.getSeconds() + 1);
    auth.service.accessTokens[resource] = {
      expiresOn: now.toISOString(),
      value: 'abc'
    }
    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert.equal(accessToken, auth.service.accessTokens[resource].value);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves new access token using existing refresh token', (done) => {
    auth.service.refreshToken = refreshToken;
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, undefined, { accessToken: 'abc' });
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      try {
        assert.equal(accessToken, 'abc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves new access token using existing refresh token (debug)', (done) => {
    auth.service.refreshToken = refreshToken;
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, undefined, { accessToken: 'abc' });
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert.equal(accessToken, 'abc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when retrieving new access token using existing refresh token', (done) => {
    auth.service.refreshToken = refreshToken;
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, { message: 'An error has occurred' }, undefined);

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving new access token using existing refresh token (debug)', (done) => {
    auth.service.refreshToken = refreshToken;
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, { message: 'An error has occurred' }, undefined);

    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows AAD error when retrieving new access token using existing refresh token', (done) => {
    auth.service.refreshToken = refreshToken;
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, { message: 'An error has occurred' }, { error_description: 'AADSTS00000 An error has occurred' });

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'AADSTS00000 An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves new access token using existing refresh token when the access token expired (debug)', (done) => {
    const now = new Date();
    now.setSeconds(now.getSeconds() - 1);
    auth.service.accessTokens[resource] = {
      expiresOn: now.toISOString(),
      value: 'abc'
    }
    auth.service.refreshToken = refreshToken;
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, undefined, { accessToken: 'acc' });
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert.equal(accessToken, 'acc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('starts device code authentication flow when no refresh token available and no authType specified', (done) => {
    const acquireUserCodeStub = sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      try {
        assert(acquireUserCodeStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when obtaining device code failed', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, { message: 'An error has occurred' }, undefined);

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows AAD error when obtaining device code failed', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, { message: 'An error has occurred' }, { error_description: 'AADSTS00000 An error has occurred' });

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'AADSTS00000 An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles device code auth error', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, { message: 'An error has occurred' }, undefined);

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles device code auth error (debug)', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, { message: 'An error has occurred' }, undefined);

    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows AAD error when device code auth fails', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, { message: 'An error has occurred' }, { error_description: 'AADSTS00000 An error has occurred' });

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'AADSTS00000 An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves access token after device code auth completed', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, { accessToken: 'abc' });
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      try {
        assert.equal(accessToken, 'abc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves access token after device code auth completed (debug)', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, { accessToken: 'abc' });
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert.equal(accessToken, 'abc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using device code authentication flow when authType deviceCode specified', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    const acquireTokenWithDeviceCodeStub = sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.service.authType = AuthType.DeviceCode;
    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      try {
        assert(acquireTokenWithDeviceCodeStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('cancels active device code authentication flow when authentication cancelled', () => {
    (auth as any).userCodeInfo = {};
    const cancelRequestToGetTokenWithDeviceCodeStub = sinon.stub((auth as any).authCtx, 'cancelRequestToGetTokenWithDeviceCode').callsFake(() => { });

    auth.cancel();
    assert(cancelRequestToGetTokenWithDeviceCodeStub.called);
  });

  it('doesn\'t cancel device code authentication flow when authentication cancelled but no flow active', () => {
    (auth as any).userCodeInfo = undefined;
    const cancelRequestToGetTokenWithDeviceCodeStub = sinon.stub((auth as any).authCtx, 'cancelRequestToGetTokenWithDeviceCode').callsFake(() => { });

    auth.cancel();
    assert(cancelRequestToGetTokenWithDeviceCodeStub.notCalled);
  });

  it('retrieves token using password flow when authType password specified', (done) => {
    const acquireTokenWithUsernamePassword = sinon.stub((auth as any).authCtx, 'acquireTokenWithUsernamePassword').callsArgWith(4, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.service.authType = AuthType.Password;
    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      try {
        assert(acquireTokenWithUsernamePassword.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using password flow when authType password specified (debug)', (done) => {
    const acquireTokenWithUsernamePassword = sinon.stub((auth as any).authCtx, 'acquireTokenWithUsernamePassword').callsArgWith(4, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.service.authType = AuthType.Password;
    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert(acquireTokenWithUsernamePassword.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when retrieving token using password flow failed', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithUsernamePassword').callsArgWith(4, { message: 'An error has occurred' }, undefined);

    auth.service.authType = AuthType.Password;
    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs error when retrieving token using password flow failed in debug mode', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithUsernamePassword').callsArgWith(4, { message: 'An error has occurred' }, { error_description: 'An error has occurred' });

    auth.service.authType = AuthType.Password;
    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert(stdoutLogSpy.calledWith({ error_description: 'An error has occurred' }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves token using certificate flow when authType certificate specified ', (done) => {
    const ensureAccessTokenWithCertificate = sinon.stub((auth as any).authCtx, 'acquireTokenWithClientCertificate').callsArgWith(4, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.service.authType = AuthType.Certificate;
    auth.ensureAccessToken(resource, stdout, false).then((accessToken) => {
      try {
        assert(ensureAccessTokenWithCertificate.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using certificate flow when authType certificate specified (debug)', (done) => {
    const ensureAccessTokenWithCertificate = sinon.stub((auth as any).authCtx, 'acquireTokenWithClientCertificate').callsArgWith(4, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.service.authType = AuthType.Certificate;
    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert(ensureAccessTokenWithCertificate.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when retrieving token using certificate flow failed', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithClientCertificate').callsArgWith(4, { message: 'An error has occurred' }, undefined);

    auth.service.authType = AuthType.Certificate;
    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert.equal(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs error when retrieving token using certificate flow failed in debug mode', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithClientCertificate').callsArgWith(4, { message: 'An error has occurred' }, { error_description: 'An error has occurred' });

    auth.service.authType = AuthType.Certificate;
    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert(stdoutLogSpy.calledWith({ error_description: 'An error has occurred' }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns access token if persisting connection fails', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done();
    }, (err) => {
      done(err);
    });
  });

  it('logs error message if persisting connection fails in debug mode', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));

    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert(stdoutLogSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('configures KeychainTokenStorage as token storage when OS is macOS', (done) => {
    sinon.stub(os, 'platform').callsFake(() => 'darwin');
    const actual = auth.getTokenStorage();
    try {
      assert(actual instanceof KeychainTokenStorage);
      done();
    }
    catch (e) {
      done(e);
    }
    finally {
      Utils.restore(os.platform);
    }
  });

  it('configures WindowsTokenStorage as token storage when OS is Windows', (done) => {
    sinon.stub(os, 'platform').callsFake(() => 'win32');
    const actual = auth.getTokenStorage();
    try {
      assert(actual instanceof WindowsTokenStorage);
      done();
    }
    catch (e) {
      done(e);
    }
    finally {
      Utils.restore(os.platform);
    }
  });

  it('configures FileTokenStorage as token storage when OS is Linux', (done) => {
    sinon.stub(os, 'platform').callsFake(() => 'linux');
    const actual = auth.getTokenStorage();
    try {
      assert(actual instanceof FileTokenStorage);
      done();
    }
    catch (e) {
      done(e);
    }
    finally {
      Utils.restore(os.platform);
    }
  });

  it('restores authentication', (done) => {
    const service: Service = new Service();
    service.refreshToken = 'abc';
    const mockStorage = {
      get: () => Promise.resolve(JSON.stringify(service))
    };
    sinon.stub(auth, 'getTokenStorage').callsFake(() => mockStorage);

    auth
      .restoreAuth()
      .then(() => {
        try {
          assert.equal(auth.service.refreshToken, 'abc');
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err) => {
        done(err);
      });
  });

  it('handles error when restoring authentication', (done) => {
    sinon.stub(auth as any, 'getServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));

    auth
      .restoreAuth()
      .then(() => {
        try {
          assert.equal(auth.service.connected, false);
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err) => {
        done(err);
      });
  });

  it('doesn\'t fail when restoring authentication from an incorrect JSON string', (done) => {
    const service: Service = new Service();
    service.refreshToken = 'abc';
    const mockStorage = {
      get: () => Promise.resolve('abc')
    };
    sinon.stub(auth, 'getTokenStorage').callsFake(() => mockStorage);

    auth
      .restoreAuth()
      .then(() => {
        assert.strictEqual(auth.service.connected, false);
        done();
      }, (err) => {
        done(err);
      });
  });

  it('doesn\'t fail when restoring authentication failed', (done) => {
    const service: Service = new Service();
    service.refreshToken = 'abc';
    const mockStorage = {
      get: () => Promise.reject('abc')
    };
    sinon.stub(auth, 'getTokenStorage').callsFake(() => mockStorage);

    auth
      .restoreAuth()
      .then(() => {
        assert.strictEqual(auth.service.connected, false);
        done();
      }, (err) => {
        done(err);
      });
  });

  it('stores connection information in the configured token storage', (done) => {
    const mockStorage = new MockTokenStorage();
    const mockStorageSetStub = sinon.stub(mockStorage, 'set').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getTokenStorage').callsFake(() => mockStorage);

    auth
      .storeConnectionInfo()
      .then(() => {
        try {
          assert(mockStorageSetStub.called);
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('clears connection information in the configured token storage', (done) => {
    const mockStorage = new MockTokenStorage();
    const mockStorageRemoveStub = sinon.stub(mockStorage, 'remove').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getTokenStorage').callsFake(() => mockStorage);

    auth
      .clearConnectionInfo()
      .then(() => {
        try {
          assert(mockStorageRemoveStub.called);
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('resets connection information on logout', () => {
    auth.service.connected = true;
    auth.service.accessTokens[resource] = {
      expiresOn: new Date().toISOString(),
      value: 'abc'
    };
    auth.service.refreshToken = 'ref';
    auth.service.authType = AuthType.Certificate;
    auth.service.userName = 'user';
    auth.service.password = 'pwd';
    auth.service.certificate = 'cert';
    auth.service.thumbprint = 'thumb';
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    auth.service.tenantId = '123';

    auth.service.logout();

    assert.strictEqual(auth.service.connected, false, 'connected');
    assert.strictEqual(JSON.stringify(auth.service.accessTokens), JSON.stringify({}), 'accessTokens');
    assert.strictEqual(auth.service.refreshToken, undefined, 'refreshToken');
    assert.strictEqual(auth.service.authType, AuthType.DeviceCode, 'authType');
    assert.strictEqual(auth.service.userName, undefined, 'userName');
    assert.strictEqual(auth.service.password, undefined, 'password');
    assert.strictEqual(auth.service.certificate, undefined, 'certificate');
    assert.strictEqual(auth.service.thumbprint, undefined, 'thumbprint');
    assert.strictEqual(auth.service.spoUrl, undefined, 'spoUrl');
    assert.strictEqual(auth.service.tenantId, undefined, 'tenantId');
  });

  it('uses the Microsoft Graph to authenticate', () => {
    assert.strictEqual(auth.defaultResource, 'https://graph.microsoft.com');
  });

  it('correctly retrieves resource from the root SharePoint site URL without trailing slash', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://contoso.sharepoint.com'), 'https://contoso.sharepoint.com');
  });

  it('correctly retrieves resource from the root SharePoint site URL with trailing slash', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://contoso.sharepoint.com/'), 'https://contoso.sharepoint.com');
  });

  it('correctly retrieves resource from a SharePoint subsite', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://contoso.sharepoint.com/subsite'), 'https://contoso.sharepoint.com');
  });

  it('correctly retrieves resource from a SharePoint site collection', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://contoso.sharepoint.com/sites/team-a'), 'https://contoso.sharepoint.com');
  });
});