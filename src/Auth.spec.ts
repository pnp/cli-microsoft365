import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
import * as request from 'request-promise-native';
import Auth, { Service } from './Auth';

class MockService extends Service {
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

  it('returns existing access token if still valid (verbose)', (done) => {
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

  it('retrieves new access token using existing refresh token (verbose)', (done) => {
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

  it('handles error when retrieving new access token using existing refresh token (verbose)', (done) => {
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

  it('waits if device code auth is still pending (verbose)', (done) => {
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

  it('correctly handles device code auth error (verbose)', (done) => {
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

  it('retrieves access token after device code auth completed (verbose)', (done) => {
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

  it('retrieves access token using refresh token (verbose)', (done) => {
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
});