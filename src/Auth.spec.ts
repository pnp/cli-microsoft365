import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
import { Auth, Service, AuthType } from './Auth';
import * as os from 'os';
import { KeychainTokenStorage } from './auth/KeychainTokenStorage';
import { WindowsTokenStorage } from './auth/WindowsTokenStorage';
import { FileTokenStorage } from './auth/FileTokenStorage';
import { TokenStorage } from './auth/TokenStorage';
import { fail } from 'assert';
import { CommandError } from './Command';

class MockService extends Service {
}

class MockAuth extends Auth {
  protected serviceId(): string {
    return 'mock';
  }
  public getConnectionInfo(): Promise<MockService> {
    return this.getServiceConnectionInfo('mock');
  }
}

class MockTokenStorage implements TokenStorage {
  public get(service: string): Promise<string> {
    return Promise.resolve('ABC');
  }

  public set(service: string): Promise<void> {
    return Promise.resolve();
  }

  public remove(service: string): Promise<void> {
    return Promise.resolve();
  }
}

describe('Auth', () => {
  let log: any[];
  let auth: Auth;
  let service: MockService;
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
    service = new MockService();
    auth = new MockAuth(service, appId);
  });

  afterEach(() => {
  });

  it('returns existing access token if still valid', (done) => {
    const now = new Date();
    now.setSeconds(now.getSeconds() + 1);
    service.accessToken = 'abc';
    service.expiresOn = now.toISOString();
    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      try {
        assert.equal(accessToken, service.accessToken);
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
    service.accessToken = 'abc';
    service.expiresOn = now.toISOString();
    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert.equal(accessToken, service.accessToken);
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
    service.refreshToken = refreshToken;
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, undefined, { accessToken: service.accessToken });
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      try {
        assert.equal(accessToken, service.accessToken);
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
    service.refreshToken = refreshToken;
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, undefined, { accessToken: service.accessToken });
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());

    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert.equal(accessToken, service.accessToken);
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
    service.refreshToken = refreshToken;
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
    service.refreshToken = refreshToken;
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
    service.refreshToken = refreshToken;
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

  it('starts device code authentication flow when no refresh token available and no authType specified', (done) => {
    const acquireUserCodeStub = sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, {});
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());

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
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());

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
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());

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
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());

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
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());

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
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.resolve());

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

  it('returns access token if persisting connection fails', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, {});
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));

    auth.ensureAccessToken(resource, stdout).then((accessToken) => {
      done();
    }, (err) => {
      done(err);
    });
  });

  it('logs error message if persisting connection fails in debug mode', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireUserCode').callsArgWith(3, undefined, {});
    sinon.stub((auth as any).authCtx, 'acquireTokenWithDeviceCode').callsArgWith(3, undefined, {});
    sinon.stub(auth as any, 'setServiceConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));

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

  it('gets access token using refresh token for the specified resource', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, undefined, { accessToken: 'acc' });

    auth.getAccessToken(resource, 'ref', stdout).then((accessToken) => {
      try {
        assert(accessToken, 'abc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('gets access token using refresh token for the specified resource (debug)', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, undefined, { accessToken: 'acc' });

    auth.getAccessToken(resource, 'ref', stdout, true).then((accessToken) => {
      try {
        assert(accessToken, 'abc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when getting access token using refresh token for the specified resource', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, { message: 'An error has occurred' }, undefined);

    auth.getAccessToken(resource, 'ref', stdout).then((accessToken) => {
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

  it('shows AAD error when getting access token using refresh token for the specified resource', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, { message: 'An error has occurred' }, { error_description: 'AADSTS00000 An error has occurred' });

    auth.getAccessToken(resource, 'ref', stdout).then((accessToken) => {
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

  it('logs the error message when getting access token using refresh token for the specified resource failed in debug mode', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithRefreshToken').callsArgWith(3, { message: 'An error has occurred' }, undefined);

    auth.getAccessToken(resource, 'ref', stdout, true).then((accessToken) => {
      done('Got access token');
    }, (err) => {
      try {
        assert(stdoutLogSpy.calledWith({ message: 'An error has occurred' }));
        done();
      }
      catch (e) {
        done(e);
      }
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
    const mockStorage = {
      get: () => Promise.resolve(JSON.stringify({ resource: 'mock' }))
    };
    sinon.stub(auth, 'getTokenStorage').callsFake(() => mockStorage);

    auth
      .restoreAuth()
      .then(() => {
        try {
          assert.equal(auth.service.resource, 'mock');
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

  it('retrieves connection information from the configured token storage', (done) => {
    const mockStorage = new MockTokenStorage();
    sinon.stub(mockStorage, 'get').callsFake(() => Promise.resolve(JSON.stringify({
      connected: true,
      resource: 'https://contoso.sharepoint.com'
    })));
    const mockAuth = new MockAuth(new MockService());
    sinon.stub(mockAuth, 'getTokenStorage').callsFake(() => mockStorage);

    mockAuth
      .getConnectionInfo()
      .then((service: MockService) => {
        try {
          assert.equal(service.connected, true);
          assert.equal(service.resource, 'https://contoso.sharepoint.com');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('correctly handles error when retrieving connection information from the configured token storage', (done) => {
    const mockStorage = new MockTokenStorage();
    sinon.stub(mockStorage, 'get').callsFake(() => Promise.reject('An error has occurred'));
    const mockAuth = new MockAuth(new MockService());
    sinon.stub(mockAuth, 'getTokenStorage').callsFake(() => mockStorage);

    mockAuth
      .getConnectionInfo()
      .then((service: MockService) => {
        fail('Expected failure but passed');
      }, (error: any) => {
        try {
          assert.equal(error, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('stores connection information in the configured token storage', (done) => {
    const mockStorage = new MockTokenStorage();
    const mockStorageSetStub = sinon.stub(mockStorage, 'set').callsFake(() => Promise.resolve());
    const mockAuth = new MockAuth(new MockService());
    sinon.stub(mockAuth, 'getTokenStorage').callsFake(() => mockStorage);

    mockAuth
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
    const mockAuth = new MockAuth(new MockService());
    sinon.stub(mockAuth, 'getTokenStorage').callsFake(() => mockStorage);

    mockAuth
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

  it('stores connection information for the specified service', (done) => {
    const mockService = new MockService();
    const mockAuth = new MockAuth(mockService);
    const setServiceConnectionInfoStub = sinon.stub(mockAuth as any, 'setServiceConnectionInfo').callsFake((serviceId, service) => Promise.resolve());

    mockAuth
      .storeConnectionInfo()
      .then(() => {
        try {
          assert(setServiceConnectionInfoStub.calledWith((mockAuth as any).serviceId(), mockService));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });
});