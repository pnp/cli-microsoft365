import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
import * as request from 'request-promise-native';
import Auth, { Service } from './Auth';
import * as os from 'os';
import { KeychainTokenStorage } from './auth/KeychainTokenStorage';
import { WindowsTokenStorage } from './auth/WindowsTokenStorage';
import { FileTokenStorage } from './auth/FileTokenStorage';
import { TokenStorage } from './auth/TokenStorage';
import { fail } from 'assert';

class MockService extends Service {
}

class MockAuth extends Auth {
  public getConnectionInfo(): Promise<MockService> {
    return this.getServiceConnectionInfo('mock');
  }

  public setConnectionInfo(): Promise<void> {
    return this.setServiceConnectionInfo('mock', {});
  }

  public clearConnectionInfo(): Promise<void> {
    return this.clearServiceConnectionInfo('mock');    
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
  let log: string[];
  let auth: Auth;
  let service: MockService;
  const resource: string = 'https://contoso.sharepoint.com';
  const appId: string = '9bc3ab49-b65d-410a-85ad-de819febfddc';
  const deviceCode: string = 'GAAABAAEAiL9Kn2Z27Uu';
  const refreshToken: string = 'ref';
  const stdout: any = {
    log: (msg: string) => {
      log.push(msg);
    }
  }
  let requestGetStub: sinon.SinonStub;

  before(() => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=refresh_token&refresh_token=${refreshToken}`) {
        return Promise.resolve('acc');
      }

      return Promise.reject('Invalid request');
    });
    requestGetStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://login.microsoftonline.com/common/oauth2/devicecode?resource=${resource}&client_id=${appId}` &&
        opts.headers.accept === 'application/json') {
        return Promise.resolve({
          interval: 5,
          device_code: deviceCode,
          message: 'To sign in, use a web browser to open the page https://aka.ms/devicelogin. Enter the code GXGPCE4CC to authenticate.'
        });
      }

      return Promise.reject('Invalid request');
    });
  });

  beforeEach(() => {
    log = [];
    service = new MockService();
    auth = new Auth(service, appId);
  });

  afterEach(() => {
  });

  after(() => {
    Utils.restore([
      request.post,
      request.get
    ]);
  });

  it('returns existing access token if still valid', (done) => {
    const now = new Date().getTime() / 1000;
    service.accessToken = 'abc';
    service.expiresAt = now + 1000;
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
    const now = new Date().getTime() / 1000;
    service.accessToken = 'abc';
    service.expiresAt = now + 1000;
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
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=refresh_token&refresh_token=${refreshToken}`) {
        return Promise.reject({
          error: {
            "error": "error",
            "error_description": "Error\r\nTrace ID: 14613dff-d719-4b49-a937-b623263415a9\r\nCorrelation ID: f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06\r\nTimestamp: 2016-03-12 01:18:44Z",
            "error_codes": [
              70016
            ],
            "timestamp": "2016-03-12 01:18:44Z",
            "trace_id": "14613dff-d719-4b49-a937-b623263415a9",
            "correlation_id": "f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06"
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    service.refreshToken = refreshToken;
    auth.ensureAccessToken(resource, stdout)
      .then((accessToken) => {
        done('Got access token');
      }, (err: any) => {
        done();
      });
  });

  it('handles error when retrieving new access token using existing refresh token (debug)', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=refresh_token&refresh_token=${refreshToken}`) {
        return Promise.reject({
          error: {
            "error": "error",
            "error_description": "Error\r\nTrace ID: 14613dff-d719-4b49-a937-b623263415a9\r\nCorrelation ID: f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06\r\nTimestamp: 2016-03-12 01:18:44Z",
            "error_codes": [
              70016
            ],
            "timestamp": "2016-03-12 01:18:44Z",
            "trace_id": "14613dff-d719-4b49-a937-b623263415a9",
            "correlation_id": "f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06"
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    service.refreshToken = refreshToken;
    auth.ensureAccessToken(resource, stdout, true)
      .then((accessToken) => {
        done('Got access token');
      }, (err: any) => {
        done();
      });
  });

  it('starts device code authentication flow when no refresh token available', (done) => {
    sinon.stub(global, 'setInterval').callsFake(() => {
      try {
        assert(requestGetStub.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(setInterval);
      }
    });
    auth.ensureAccessToken(resource, stdout, true);
  });

  it('checks if the device code auth completed on the given interval', (done) => {
    sinon.stub(global, 'setInterval').callsFake((cb, interval) => {
      try {
        assert.equal(interval, 5000);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(setInterval);
      }
    });
    auth.ensureAccessToken(resource, stdout, true);
  });

  it('waits if device code auth is still pending', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=device_code&code=${deviceCode}`) {
        return Promise.reject({
          error: {
            "error": "authorization_pending",
            "error_description": "AADSTS70016: Pending end-user authorization.\r\nTrace ID: 14613dff-d719-4b49-a937-b623263415a9\r\nCorrelation ID: f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06\r\nTimestamp: 2016-03-12 01:18:44Z",
            "error_codes": [
              70016
            ],
            "timestamp": "2016-03-12 01:18:44Z",
            "trace_id": "14613dff-d719-4b49-a937-b623263415a9",
            "correlation_id": "f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06"
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setInterval').callsFake((cb) => {
      cb();
      // wait for the promise inside the callback to continue
      setTimeout(() => {
        let isWaiting = log.length === 1 && log[0] === 'To sign in, use a web browser to open the page https://aka.ms/devicelogin. Enter the code GXGPCE4CC to authenticate.';

        try {
          assert(isWaiting);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(setInterval);
        }
      }, 1);
    });
    auth.ensureAccessToken(resource, stdout);
  });

  it('waits if device code auth is still pending (debug)', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=device_code&code=${deviceCode}`) {
        return Promise.reject({
          error: {
            "error": "authorization_pending",
            "error_description": "AADSTS70016: Pending end-user authorization.\r\nTrace ID: 14613dff-d719-4b49-a937-b623263415a9\r\nCorrelation ID: f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06\r\nTimestamp: 2016-03-12 01:18:44Z",
            "error_codes": [
              70016
            ],
            "timestamp": "2016-03-12 01:18:44Z",
            "trace_id": "14613dff-d719-4b49-a937-b623263415a9",
            "correlation_id": "f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06"
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setInterval').callsFake((cb) => {
      cb();
      // wait for the promise inside the callback to continue
      setTimeout(() => {
        let isWaiting = false;
        log.forEach(l => {
          if (l && l === 'Authorization pending...') {
            isWaiting = true;
          }
        });

        try {
          assert(isWaiting);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(setInterval);
        }
      }, 1);
    });
    auth.ensureAccessToken(resource, stdout, true);
  });

  it('correctly handles device code auth error', (done) => {
    Utils.restore([
      request.post,
      setInterval
    ]);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=device_code&code=${deviceCode}`) {
        return Promise.reject({
          error: {
            "error": "error",
            "error_description": "Error\r\nTrace ID: 14613dff-d719-4b49-a937-b623263415a9\r\nCorrelation ID: f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06\r\nTimestamp: 2016-03-12 01:18:44Z",
            "error_codes": [
              70016
            ],
            "timestamp": "2016-03-12 01:18:44Z",
            "trace_id": "14613dff-d719-4b49-a937-b623263415a9",
            "correlation_id": "f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06"
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setInterval').callsFake((cb) => {
      cb();
    });
    auth.ensureAccessToken(resource, stdout)
      .then((accessToken) => {
        done('Got access token');
      }, (err: any) => {
        if (err === 'error') {
          done();
        }
        else {
          done(err);
        }
      });
  });

  it('correctly handles device code auth error (debug)', (done) => {
    Utils.restore([
      request.post,
      setInterval
    ]);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=device_code&code=${deviceCode}`) {
        return Promise.reject({
          error: {
            "error": "error",
            "error_description": "Error\r\nTrace ID: 14613dff-d719-4b49-a937-b623263415a9\r\nCorrelation ID: f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06\r\nTimestamp: 2016-03-12 01:18:44Z",
            "error_codes": [
              70016
            ],
            "timestamp": "2016-03-12 01:18:44Z",
            "trace_id": "14613dff-d719-4b49-a937-b623263415a9",
            "correlation_id": "f4e1c3a8-15a8-4ae4-8389-19a5a0ce2e06"
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setInterval').callsFake((cb) => {
      cb();
    });
    auth.ensureAccessToken(resource, stdout, true)
      .then((accessToken) => {
        done('Got access token');
      }, (err: any) => {
        if (err === 'error') {
          done();
        }
        else {
          done(err);
        }
      });
  });

  it('retrieves access token after device code auth completed', (done) => {
    Utils.restore([
      request.post,
      setInterval
    ]);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=device_code&code=${deviceCode}`) {
        return Promise.resolve({
          access_token: 'acc',
          refresh_token: 'ref',
          expires_on: 0
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setInterval').callsFake((cb) => {
      cb();
    });
    auth.ensureAccessToken(resource, stdout)
      .then((accessToken) => {
        try {
          assert.equal(accessToken, 'acc');
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err: any) => {
        done(err);
      });
  });

  it('retrieves access token after device code auth completed (debug)', (done) => {
    Utils.restore([
      request.post,
      setInterval
    ]);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=device_code&code=${deviceCode}`) {
        return Promise.resolve({
          access_token: 'acc',
          refresh_token: 'ref',
          expires_on: 0
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setInterval').callsFake((cb) => {
      cb();
    });
    auth.ensureAccessToken(resource, stdout, true)
      .then((accessToken) => {
        try {
          assert.equal(accessToken, 'acc');
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err: any) => {
        done(err);
      });
  });

  it('retrieves access token using refresh token', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=refresh_token&refresh_token=${refreshToken}`) {
        return Promise.resolve({
          access_token: 'acc',
          refresh_token: 'ref',
          expires_on: 0
        });
      }

      return Promise.reject('Invalid request');
    });
    auth.getAccessToken(resource, refreshToken, stdout).then((accessToken) => {
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

  it('retrieves access token using refresh token (debug)', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=refresh_token&refresh_token=${refreshToken}`) {
        return Promise.resolve({
          access_token: 'acc',
          refresh_token: 'ref',
          expires_on: 0
        });
      }

      return Promise.reject('Invalid request');
    });
    auth.getAccessToken(resource, refreshToken, stdout, true).then((accessToken) => {
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

  it('handles error when retrieving access token using refresh token', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://login.microsoftonline.com/common/oauth2/token' &&
        opts.headers['Content-Type'] === 'application/x-www-form-urlencoded' &&
        opts.headers.accept === 'application/json' &&
        opts.body === `resource=${encodeURIComponent(resource)}&client_id=${appId}&grant_type=refresh_token&refresh_token=${refreshToken}`) {
        return Promise.reject('error');
      }

      return Promise.reject('Invalid request');
    });
    auth.getAccessToken(resource, refreshToken, stdout, true).then((accessToken) => {
      done('Retrieved access token');
    }, (err) => {
      if (err === 'error') {
        done();
      }
      else {
        done(err);
      }
    });
  });

  it('handles error when starting device code authentication flow when no refresh token available', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://login.microsoftonline.com/common/oauth2/devicecode?resource=${resource}&client_id=${appId}` &&
        opts.headers.accept === 'application/json') {
        return Promise.reject('error');
      }

      return Promise.reject('Invalid request');
    });
    auth.ensureAccessToken(resource, stdout, true)
      .then((response) => {
        done('Resolved promise despite error');
      }, (err) => {
        if (err === 'error') {
          done();
        }
        else {
          done(err);
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
    auth
      .restoreAuth()
      .then(() => {
        done();
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
    const mockStorageGetStub = sinon.stub(mockStorage, 'set').callsFake(() => Promise.resolve());
    const mockAuth = new MockAuth(new MockService());
    sinon.stub(mockAuth, 'getTokenStorage').callsFake(() => mockStorage);

    mockAuth
      .setConnectionInfo()
      .then(() => {
        try {
          assert(mockStorageGetStub.called);
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
});