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
import * as fs from 'fs';

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
  let readFileSyncStub: sinon.SinonStub;

  beforeEach(() => {
    log = [];
    auth = new Auth();
    (auth as any).appId = appId;
    readFileSyncStub = sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');
  });

  afterEach(() => {
    readFileSyncStub.restore();
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
    // base64 encoded PEM Cert
    auth.service.certificate = 'QmFnIEF0dHJpYnV0ZXMNCiAgICBsb2NhbEtleUlEOiBDQyBGNCBGMiBBMyBDMyBEMiAwOSBDNSAxMiBCMyA3MiA0QiBCOCA4MyBBNSA0NyA0QyAwOSAyMSBEQyANCnN1YmplY3Q9QyA9IEFVLCBTVCA9IFNvbWUtU3RhdGUsIE8gPSBJbnRlcm5ldCBXaWRnaXRzIFB0eSBMdGQNCg0KaXNzdWVyPUMgPSBBVSwgU1QgPSBTb21lLVN0YXRlLCBPID0gSW50ZXJuZXQgV2lkZ2l0cyBQdHkgTHRkDQoNCi0tLS0tQkVHSU4gQ0VSVElGSUNBVEUtLS0tLQ0KTUlJRGF6Q0NBbE9nQXdJQkFnSVVXb25VNFM0RTcxRjVZMU5zU0xYbUlhZ1dkNVl3RFFZSktvWklodmNOQVFFTA0KQlFBd1JURUxNQWtHQTFVRUJoTUNRVlV4RXpBUkJnTlZCQWdNQ2xOdmJXVXRVM1JoZEdVeElUQWZCZ05WQkFvTQ0KR0VsdWRHVnlibVYwSUZkcFpHZHBkSE1nVUhSNUlFeDBaREFlRncweE9UQTNNVEl5TVRVek1qbGFGdzB5TURBMw0KTVRFeU1UVXpNamxhTUVVeEN6QUpCZ05WQkFZVEFrRlZNUk13RVFZRFZRUUlEQXBUYjIxbExWTjBZWFJsTVNFdw0KSHdZRFZRUUtEQmhKYm5SbGNtNWxkQ0JYYVdSbmFYUnpJRkIwZVNCTWRHUXdnZ0VpTUEwR0NTcUdTSWIzRFFFQg0KQVFVQUE0SUJEd0F3Z2dFS0FvSUJBUUNsa01lQXlKbTJkMy95aEV0NHZGYjYrYjEyUGxRSDB4VGx1a1BoK2xScg0KOXJDNk5DM3dObnoySm5vbE1HclhuZVp2TlN5czFONVpSTm0yTjhQdy9QOExxeHJSenFFOFBNVC96NnN1UFhSUg0KWm5hZ2xaUklXb0NNR25pRVlDZVJHZnI4R2JpUXcwYlZEeXFuSnJaZjByS0pHbnZUNlY3QmpUdFloRWIzeXhoNA0KSmNUSnIrVDl0OEFYaldmemt6alBZdklxYmhha3FxcHd1SEVPYkh4T201cHVERTFBNVJOZm8wamcrTmZtVko5VQ0KMWR1RjVzdmE2NVQ5Q1RtdEdlbVNlUGlzWmgxZmhoOS94QmJwTCs0RUJWUXZqdEZXWk5zMVJHMW9QUllscmpzaQ0KTXFsaHNUdjhDZXI5cWUxcVNTdHFjMmJsc3hGek1zNmxZOHAvUHIrYm5uR3pBZ01CQUFHalV6QlJNQjBHQTFVZA0KRGdRV0JCU203cWFreXQwY2xxN0lnRFRWdkUrWEpaNFU5akFmQmdOVkhTTUVHREFXZ0JTbTdxYWt5dDBjbHE3SQ0KZ0RUVnZFK1hKWjRVOWpBUEJnTlZIUk1CQWY4RUJUQURBUUgvTUEwR0NTcUdTSWIzRFFFQkN3VUFBNElCQVFBYQ0KQnVqTytveU0yL0Q0SzNpS3lqVDVzbHF2UFVlVzFrZVVXYVdSVDZXRTY0VkFPbTlPZzU1bkIyOE5TSVVXampXMA0KdTJEUHF3SzJiOEFXalEveWp3S3NUMXVTdzcyQ0VEY2o3SkE1VXA5UWpBa0hIZmFoQWtOd0o5M0llcmFBdTEyVQ0KN25FRDdIN20yeGZscDVwM0dadzNHUE0rZmpBaDZLOUZIRDI0bWdGUTh4b2JPQSttVEVvV2ZIVVQrZ1pUMGxYdQ0KazFrVTJVelVOd2dwc3c4V04wNFFzWU5XcFF5d3ppUWtuZTQzNW5tdmxZOGZRc2hPSnErK0JCS0thd0xEcjk3bA0KRTBYQUxEZDZlVVhQenZ5OU1xZlozeUswRmUzMy8zbnZnUnE4QWZ3azRsbzhac2ZYWUlSTXA3b3BER0VmaUZmNQ0KM3JTTGxSZG9TNDQ4OVFZRnAyYUQNCi0tLS0tRU5EIENFUlRJRklDQVRFLS0tLS0NCkJhZyBBdHRyaWJ1dGVzDQogICAgbG9jYWxLZXlJRDogQ0MgRjQgRjIgQTMgQzMgRDIgMDkgQzUgMTIgQjMgNzIgNEIgQjggODMgQTUgNDcgNEMgMDkgMjEgREMgDQpLZXkgQXR0cmlidXRlczogPE5vIEF0dHJpYnV0ZXM+DQotLS0tLUJFR0lOIFBSSVZBVEUgS0VZLS0tLS0NCk1JSUV2Z0lCQURBTkJna3Foa2lHOXcwQkFRRUZBQVNDQktnd2dnU2tBZ0VBQW9JQkFRQ2xrTWVBeUptMmQzL3kNCmhFdDR2RmI2K2IxMlBsUUgweFRsdWtQaCtsUnI5ckM2TkMzd05uejJKbm9sTUdyWG5lWnZOU3lzMU41WlJObTINCk44UHcvUDhMcXhyUnpxRThQTVQvejZzdVBYUlJabmFnbFpSSVdvQ01HbmlFWUNlUkdmcjhHYmlRdzBiVkR5cW4NCkpyWmYwcktKR252VDZWN0JqVHRZaEViM3l4aDRKY1RKcitUOXQ4QVhqV2Z6a3pqUFl2SXFiaGFrcXFwd3VIRU8NCmJIeE9tNXB1REUxQTVSTmZvMGpnK05mbVZKOVUxZHVGNXN2YTY1VDlDVG10R2VtU2VQaXNaaDFmaGg5L3hCYnANCkwrNEVCVlF2anRGV1pOczFSRzFvUFJZbHJqc2lNcWxoc1R2OENlcjlxZTFxU1N0cWMyYmxzeEZ6TXM2bFk4cC8NClByK2Jubkd6QWdNQkFBRUNnZ0VBUjRsMytqZ3kybmxseWtiSlNXQ3ZnSCs2RWtZNkRxdHd3eFlwVUpIV09sUDcNCjVtaTNWS3htY0FFT0U5V0l4S05RTnNyV0E5TnlRMFlSZjc4MnBZRGJQcEp1NHlxUjFqSTN1SVJsWlhSZU52RzcNCjNnVGpiaVBVbVRTeTBCZXY0TzFGMmZuUEdwV1ZuR2VTT1dqcnNobWExTXlocGwyV2VMRHFiSU96R2t3aHhYOXkNClRhRFd5MjErbDFpNVNGWUZTdHdXOWlhOXRORTFTTTU4WnpQWk0yK0NDdHhQVEFBQXRJRmZXUVdTbnhodUxMenMNCjNyVDRVOGNLZzJITVBXb29rOS9peWxsa0xEVXBPanhJR2tHWXdheDVnR2xvR0xZYWVoelc5Q3hobzgvc3A4WjUNCkVNNVFvczVJSTF2K21pNHhHa0RTdW4rbDYzcDN5Nm54T3pqM1h1MzRlUUtCZ1FEUDNtRWttN2lVaTlhRUxweXYNCkIxeDFlRFR2UmEwcllZMHZUaXFrYzhyUGc0NU1uOUNWRWZqdnV3YkN4M21tTExabThqZVY3ZTFHWjZJeXYreEUNCmcxeFkrUTd0RUlCb1FwWThlemg0UVYvMXRkZkhiUzNPcGdIbHVqMGd5MWxqT2QrbkxzS2RNQWRlYVF3Uy9WK2MNCk51Sks0Y3oyQWl6UXU1dHQ4WHdoOGdvU0Z3S0JnUURMNXRjZnF0VmdMQWJmMnJQbEhBLzdNcU1sWGpqNUQ0ejkNCjZmTWlCVDdOWHlYUGx6a2pJQkxOdG9OWlBCVTFzeERFb2tiNUtyTlhLTUtIaU9nTkQ0cWtDYkdnRFk2WUdaS3cNCkg4bDlLWDBaM2pwcEp0TURvQ21yQW9hSmZTUXNreGJXSDd4VlFGVzdPVWQ0dHMxZ3FDbTBUTFVxeW9lcW1EK3INCmg3WFlaa2RxeFFLQmdBK2NpZnN2M3NyNVBhRXJ4d1MyTHRGN3Q2NElzNXJBZHRRSXNOY3RBeHhXcXdkQ01XNGcNCnJXdUR4bHcya3dKUjlWa0I4LzdFb2I5WjVTcWVrMllKMzVPbkVPSHBEVnZITkhWU1k4bFVUNXFxajR3Z3ZRSDYNCkljWlpHR0l3STRSNlFqdlNIVGVrOWNpM1p2cStJTUlndFJvZW4wQVNwYjcvZUFybnlnVGFvcnI5QW9HQkFJT3QNCllOSEhqaUtjYkJnV2NjU01tZGw4T3hXL3dvVTlRSzBkYjNGUjk5dkREWFVCVU5uWk5hdDVxVnR3VExZd0hLMFANCnEwdndBbjlRQ0VoazVvN0FzYVQ3eWFUMS9GZEhkSTZmQ0l6MnhSNTJnRHcxNFdIZkJlbTFLTk1UYU5BTWNWdjQNCmhMUjlacUFRL3BIN1k2aC9FT2VwL2ZsVGI4ZUFxT1dLTDZvL2F2R05Bb0dCQUlHc0c1VExuSmlPU044SUtGU04NCmJmK3IrNkhWL2R6MkluNjhSR255MTB0OGpwbUpPbGgrdXRncGtvOXI2Y09uWGY4VHM2SFAveTBtbDl5YXhvMlANCm52c2wwcFlseFQxQy9taXJaZWxYKzFaQTltdFpHT2RxbzZhdVZUM1drcXBpb3c2WUtzbzl2Z2RHWmRWRUxiMEINCnUvdyt4UjBvN21aSEpwVEdmS09KdE53MQ0KLS0tLS1FTkQgUFJJVkFURSBLRVktLS0tLQ0K';
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
    // base64 encoded PEM Cert
    auth.service.certificate = 'QmFnIEF0dHJpYnV0ZXMNCiAgICBsb2NhbEtleUlEOiBDQyBGNCBGMiBBMyBDMyBEMiAwOSBDNSAxMiBCMyA3MiA0QiBCOCA4MyBBNSA0NyA0QyAwOSAyMSBEQyANCnN1YmplY3Q9QyA9IEFVLCBTVCA9IFNvbWUtU3RhdGUsIE8gPSBJbnRlcm5ldCBXaWRnaXRzIFB0eSBMdGQNCg0KaXNzdWVyPUMgPSBBVSwgU1QgPSBTb21lLVN0YXRlLCBPID0gSW50ZXJuZXQgV2lkZ2l0cyBQdHkgTHRkDQoNCi0tLS0tQkVHSU4gQ0VSVElGSUNBVEUtLS0tLQ0KTUlJRGF6Q0NBbE9nQXdJQkFnSVVXb25VNFM0RTcxRjVZMU5zU0xYbUlhZ1dkNVl3RFFZSktvWklodmNOQVFFTA0KQlFBd1JURUxNQWtHQTFVRUJoTUNRVlV4RXpBUkJnTlZCQWdNQ2xOdmJXVXRVM1JoZEdVeElUQWZCZ05WQkFvTQ0KR0VsdWRHVnlibVYwSUZkcFpHZHBkSE1nVUhSNUlFeDBaREFlRncweE9UQTNNVEl5TVRVek1qbGFGdzB5TURBMw0KTVRFeU1UVXpNamxhTUVVeEN6QUpCZ05WQkFZVEFrRlZNUk13RVFZRFZRUUlEQXBUYjIxbExWTjBZWFJsTVNFdw0KSHdZRFZRUUtEQmhKYm5SbGNtNWxkQ0JYYVdSbmFYUnpJRkIwZVNCTWRHUXdnZ0VpTUEwR0NTcUdTSWIzRFFFQg0KQVFVQUE0SUJEd0F3Z2dFS0FvSUJBUUNsa01lQXlKbTJkMy95aEV0NHZGYjYrYjEyUGxRSDB4VGx1a1BoK2xScg0KOXJDNk5DM3dObnoySm5vbE1HclhuZVp2TlN5czFONVpSTm0yTjhQdy9QOExxeHJSenFFOFBNVC96NnN1UFhSUg0KWm5hZ2xaUklXb0NNR25pRVlDZVJHZnI4R2JpUXcwYlZEeXFuSnJaZjByS0pHbnZUNlY3QmpUdFloRWIzeXhoNA0KSmNUSnIrVDl0OEFYaldmemt6alBZdklxYmhha3FxcHd1SEVPYkh4T201cHVERTFBNVJOZm8wamcrTmZtVko5VQ0KMWR1RjVzdmE2NVQ5Q1RtdEdlbVNlUGlzWmgxZmhoOS94QmJwTCs0RUJWUXZqdEZXWk5zMVJHMW9QUllscmpzaQ0KTXFsaHNUdjhDZXI5cWUxcVNTdHFjMmJsc3hGek1zNmxZOHAvUHIrYm5uR3pBZ01CQUFHalV6QlJNQjBHQTFVZA0KRGdRV0JCU203cWFreXQwY2xxN0lnRFRWdkUrWEpaNFU5akFmQmdOVkhTTUVHREFXZ0JTbTdxYWt5dDBjbHE3SQ0KZ0RUVnZFK1hKWjRVOWpBUEJnTlZIUk1CQWY4RUJUQURBUUgvTUEwR0NTcUdTSWIzRFFFQkN3VUFBNElCQVFBYQ0KQnVqTytveU0yL0Q0SzNpS3lqVDVzbHF2UFVlVzFrZVVXYVdSVDZXRTY0VkFPbTlPZzU1bkIyOE5TSVVXampXMA0KdTJEUHF3SzJiOEFXalEveWp3S3NUMXVTdzcyQ0VEY2o3SkE1VXA5UWpBa0hIZmFoQWtOd0o5M0llcmFBdTEyVQ0KN25FRDdIN20yeGZscDVwM0dadzNHUE0rZmpBaDZLOUZIRDI0bWdGUTh4b2JPQSttVEVvV2ZIVVQrZ1pUMGxYdQ0KazFrVTJVelVOd2dwc3c4V04wNFFzWU5XcFF5d3ppUWtuZTQzNW5tdmxZOGZRc2hPSnErK0JCS0thd0xEcjk3bA0KRTBYQUxEZDZlVVhQenZ5OU1xZlozeUswRmUzMy8zbnZnUnE4QWZ3azRsbzhac2ZYWUlSTXA3b3BER0VmaUZmNQ0KM3JTTGxSZG9TNDQ4OVFZRnAyYUQNCi0tLS0tRU5EIENFUlRJRklDQVRFLS0tLS0NCkJhZyBBdHRyaWJ1dGVzDQogICAgbG9jYWxLZXlJRDogQ0MgRjQgRjIgQTMgQzMgRDIgMDkgQzUgMTIgQjMgNzIgNEIgQjggODMgQTUgNDcgNEMgMDkgMjEgREMgDQpLZXkgQXR0cmlidXRlczogPE5vIEF0dHJpYnV0ZXM+DQotLS0tLUJFR0lOIFBSSVZBVEUgS0VZLS0tLS0NCk1JSUV2Z0lCQURBTkJna3Foa2lHOXcwQkFRRUZBQVNDQktnd2dnU2tBZ0VBQW9JQkFRQ2xrTWVBeUptMmQzL3kNCmhFdDR2RmI2K2IxMlBsUUgweFRsdWtQaCtsUnI5ckM2TkMzd05uejJKbm9sTUdyWG5lWnZOU3lzMU41WlJObTINCk44UHcvUDhMcXhyUnpxRThQTVQvejZzdVBYUlJabmFnbFpSSVdvQ01HbmlFWUNlUkdmcjhHYmlRdzBiVkR5cW4NCkpyWmYwcktKR252VDZWN0JqVHRZaEViM3l4aDRKY1RKcitUOXQ4QVhqV2Z6a3pqUFl2SXFiaGFrcXFwd3VIRU8NCmJIeE9tNXB1REUxQTVSTmZvMGpnK05mbVZKOVUxZHVGNXN2YTY1VDlDVG10R2VtU2VQaXNaaDFmaGg5L3hCYnANCkwrNEVCVlF2anRGV1pOczFSRzFvUFJZbHJqc2lNcWxoc1R2OENlcjlxZTFxU1N0cWMyYmxzeEZ6TXM2bFk4cC8NClByK2Jubkd6QWdNQkFBRUNnZ0VBUjRsMytqZ3kybmxseWtiSlNXQ3ZnSCs2RWtZNkRxdHd3eFlwVUpIV09sUDcNCjVtaTNWS3htY0FFT0U5V0l4S05RTnNyV0E5TnlRMFlSZjc4MnBZRGJQcEp1NHlxUjFqSTN1SVJsWlhSZU52RzcNCjNnVGpiaVBVbVRTeTBCZXY0TzFGMmZuUEdwV1ZuR2VTT1dqcnNobWExTXlocGwyV2VMRHFiSU96R2t3aHhYOXkNClRhRFd5MjErbDFpNVNGWUZTdHdXOWlhOXRORTFTTTU4WnpQWk0yK0NDdHhQVEFBQXRJRmZXUVdTbnhodUxMenMNCjNyVDRVOGNLZzJITVBXb29rOS9peWxsa0xEVXBPanhJR2tHWXdheDVnR2xvR0xZYWVoelc5Q3hobzgvc3A4WjUNCkVNNVFvczVJSTF2K21pNHhHa0RTdW4rbDYzcDN5Nm54T3pqM1h1MzRlUUtCZ1FEUDNtRWttN2lVaTlhRUxweXYNCkIxeDFlRFR2UmEwcllZMHZUaXFrYzhyUGc0NU1uOUNWRWZqdnV3YkN4M21tTExabThqZVY3ZTFHWjZJeXYreEUNCmcxeFkrUTd0RUlCb1FwWThlemg0UVYvMXRkZkhiUzNPcGdIbHVqMGd5MWxqT2QrbkxzS2RNQWRlYVF3Uy9WK2MNCk51Sks0Y3oyQWl6UXU1dHQ4WHdoOGdvU0Z3S0JnUURMNXRjZnF0VmdMQWJmMnJQbEhBLzdNcU1sWGpqNUQ0ejkNCjZmTWlCVDdOWHlYUGx6a2pJQkxOdG9OWlBCVTFzeERFb2tiNUtyTlhLTUtIaU9nTkQ0cWtDYkdnRFk2WUdaS3cNCkg4bDlLWDBaM2pwcEp0TURvQ21yQW9hSmZTUXNreGJXSDd4VlFGVzdPVWQ0dHMxZ3FDbTBUTFVxeW9lcW1EK3INCmg3WFlaa2RxeFFLQmdBK2NpZnN2M3NyNVBhRXJ4d1MyTHRGN3Q2NElzNXJBZHRRSXNOY3RBeHhXcXdkQ01XNGcNCnJXdUR4bHcya3dKUjlWa0I4LzdFb2I5WjVTcWVrMllKMzVPbkVPSHBEVnZITkhWU1k4bFVUNXFxajR3Z3ZRSDYNCkljWlpHR0l3STRSNlFqdlNIVGVrOWNpM1p2cStJTUlndFJvZW4wQVNwYjcvZUFybnlnVGFvcnI5QW9HQkFJT3QNCllOSEhqaUtjYkJnV2NjU01tZGw4T3hXL3dvVTlRSzBkYjNGUjk5dkREWFVCVU5uWk5hdDVxVnR3VExZd0hLMFANCnEwdndBbjlRQ0VoazVvN0FzYVQ3eWFUMS9GZEhkSTZmQ0l6MnhSNTJnRHcxNFdIZkJlbTFLTk1UYU5BTWNWdjQNCmhMUjlacUFRL3BIN1k2aC9FT2VwL2ZsVGI4ZUFxT1dLTDZvL2F2R05Bb0dCQUlHc0c1VExuSmlPU044SUtGU04NCmJmK3IrNkhWL2R6MkluNjhSR255MTB0OGpwbUpPbGgrdXRncGtvOXI2Y09uWGY4VHM2SFAveTBtbDl5YXhvMlANCm52c2wwcFlseFQxQy9taXJaZWxYKzFaQTltdFpHT2RxbzZhdVZUM1drcXBpb3c2WUtzbzl2Z2RHWmRWRUxiMEINCnUvdyt4UjBvN21aSEpwVEdmS09KdE53MQ0KLS0tLS1FTkQgUFJJVkFURSBLRVktLS0tLQ0K';
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

  it('retrieves token using PFX certificate flow when authType certificate specified (debug)', (done) => {
    readFileSyncStub.restore();
    const ensureAccessTokenStub = sinon.stub((auth as any).authCtx, 'acquireTokenWithClientCertificate').callsArgWith(4, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.service.authType = AuthType.Certificate;
    auth.service.password = 'pass@word1';
    // base64 encoded PFX file
    auth.service.certificate = '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';
    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
      try {
        assert.equal(ensureAccessTokenStub.lastCall.args[0], "https://contoso.sharepoint.com");
        assert.equal(ensureAccessTokenStub.lastCall.args[1], "9bc3ab49-b65d-410a-85ad-de819febfddc");
        assert.notEqual(ensureAccessTokenStub.lastCall.args[2].indexOf("MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQ"), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when PFX certificate flow when authType certificate specified (debug)', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithClientCertificate').callsArgWith(4, undefined, {});
    sinon.stub(auth as any, 'storeConnectionInfo').callsFake(() => Promise.resolve());

    auth.service.authType = AuthType.Certificate;
    // base64 encoded PFX file
    auth.service.certificate = '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';
    auth.service.password = 'abc';
    auth.ensureAccessToken(resource, stdout, true).then((accessToken) => {
        done();
    }, (err) => {
      try {
        assert.equal(err, 'Error: PKCS#12 MAC could not be verified. Invalid password?');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving token using certificate flow failed', (done) => {
    sinon.stub((auth as any).authCtx, 'acquireTokenWithClientCertificate').callsArgWith(4, { message: 'An error has occurred' }, undefined);

    auth.service.authType = AuthType.Certificate;
    // base64 encoded PEM cert
    auth.service.certificate = 'QmFnIEF0dHJpYnV0ZXMNCiAgICBsb2NhbEtleUlEOiBDQyBGNCBGMiBBMyBDMyBEMiAwOSBDNSAxMiBCMyA3MiA0QiBCOCA4MyBBNSA0NyA0QyAwOSAyMSBEQyANCnN1YmplY3Q9QyA9IEFVLCBTVCA9IFNvbWUtU3RhdGUsIE8gPSBJbnRlcm5ldCBXaWRnaXRzIFB0eSBMdGQNCg0KaXNzdWVyPUMgPSBBVSwgU1QgPSBTb21lLVN0YXRlLCBPID0gSW50ZXJuZXQgV2lkZ2l0cyBQdHkgTHRkDQoNCi0tLS0tQkVHSU4gQ0VSVElGSUNBVEUtLS0tLQ0KTUlJRGF6Q0NBbE9nQXdJQkFnSVVXb25VNFM0RTcxRjVZMU5zU0xYbUlhZ1dkNVl3RFFZSktvWklodmNOQVFFTA0KQlFBd1JURUxNQWtHQTFVRUJoTUNRVlV4RXpBUkJnTlZCQWdNQ2xOdmJXVXRVM1JoZEdVeElUQWZCZ05WQkFvTQ0KR0VsdWRHVnlibVYwSUZkcFpHZHBkSE1nVUhSNUlFeDBaREFlRncweE9UQTNNVEl5TVRVek1qbGFGdzB5TURBMw0KTVRFeU1UVXpNamxhTUVVeEN6QUpCZ05WQkFZVEFrRlZNUk13RVFZRFZRUUlEQXBUYjIxbExWTjBZWFJsTVNFdw0KSHdZRFZRUUtEQmhKYm5SbGNtNWxkQ0JYYVdSbmFYUnpJRkIwZVNCTWRHUXdnZ0VpTUEwR0NTcUdTSWIzRFFFQg0KQVFVQUE0SUJEd0F3Z2dFS0FvSUJBUUNsa01lQXlKbTJkMy95aEV0NHZGYjYrYjEyUGxRSDB4VGx1a1BoK2xScg0KOXJDNk5DM3dObnoySm5vbE1HclhuZVp2TlN5czFONVpSTm0yTjhQdy9QOExxeHJSenFFOFBNVC96NnN1UFhSUg0KWm5hZ2xaUklXb0NNR25pRVlDZVJHZnI4R2JpUXcwYlZEeXFuSnJaZjByS0pHbnZUNlY3QmpUdFloRWIzeXhoNA0KSmNUSnIrVDl0OEFYaldmemt6alBZdklxYmhha3FxcHd1SEVPYkh4T201cHVERTFBNVJOZm8wamcrTmZtVko5VQ0KMWR1RjVzdmE2NVQ5Q1RtdEdlbVNlUGlzWmgxZmhoOS94QmJwTCs0RUJWUXZqdEZXWk5zMVJHMW9QUllscmpzaQ0KTXFsaHNUdjhDZXI5cWUxcVNTdHFjMmJsc3hGek1zNmxZOHAvUHIrYm5uR3pBZ01CQUFHalV6QlJNQjBHQTFVZA0KRGdRV0JCU203cWFreXQwY2xxN0lnRFRWdkUrWEpaNFU5akFmQmdOVkhTTUVHREFXZ0JTbTdxYWt5dDBjbHE3SQ0KZ0RUVnZFK1hKWjRVOWpBUEJnTlZIUk1CQWY4RUJUQURBUUgvTUEwR0NTcUdTSWIzRFFFQkN3VUFBNElCQVFBYQ0KQnVqTytveU0yL0Q0SzNpS3lqVDVzbHF2UFVlVzFrZVVXYVdSVDZXRTY0VkFPbTlPZzU1bkIyOE5TSVVXampXMA0KdTJEUHF3SzJiOEFXalEveWp3S3NUMXVTdzcyQ0VEY2o3SkE1VXA5UWpBa0hIZmFoQWtOd0o5M0llcmFBdTEyVQ0KN25FRDdIN20yeGZscDVwM0dadzNHUE0rZmpBaDZLOUZIRDI0bWdGUTh4b2JPQSttVEVvV2ZIVVQrZ1pUMGxYdQ0KazFrVTJVelVOd2dwc3c4V04wNFFzWU5XcFF5d3ppUWtuZTQzNW5tdmxZOGZRc2hPSnErK0JCS0thd0xEcjk3bA0KRTBYQUxEZDZlVVhQenZ5OU1xZlozeUswRmUzMy8zbnZnUnE4QWZ3azRsbzhac2ZYWUlSTXA3b3BER0VmaUZmNQ0KM3JTTGxSZG9TNDQ4OVFZRnAyYUQNCi0tLS0tRU5EIENFUlRJRklDQVRFLS0tLS0NCkJhZyBBdHRyaWJ1dGVzDQogICAgbG9jYWxLZXlJRDogQ0MgRjQgRjIgQTMgQzMgRDIgMDkgQzUgMTIgQjMgNzIgNEIgQjggODMgQTUgNDcgNEMgMDkgMjEgREMgDQpLZXkgQXR0cmlidXRlczogPE5vIEF0dHJpYnV0ZXM+DQotLS0tLUJFR0lOIFBSSVZBVEUgS0VZLS0tLS0NCk1JSUV2Z0lCQURBTkJna3Foa2lHOXcwQkFRRUZBQVNDQktnd2dnU2tBZ0VBQW9JQkFRQ2xrTWVBeUptMmQzL3kNCmhFdDR2RmI2K2IxMlBsUUgweFRsdWtQaCtsUnI5ckM2TkMzd05uejJKbm9sTUdyWG5lWnZOU3lzMU41WlJObTINCk44UHcvUDhMcXhyUnpxRThQTVQvejZzdVBYUlJabmFnbFpSSVdvQ01HbmlFWUNlUkdmcjhHYmlRdzBiVkR5cW4NCkpyWmYwcktKR252VDZWN0JqVHRZaEViM3l4aDRKY1RKcitUOXQ4QVhqV2Z6a3pqUFl2SXFiaGFrcXFwd3VIRU8NCmJIeE9tNXB1REUxQTVSTmZvMGpnK05mbVZKOVUxZHVGNXN2YTY1VDlDVG10R2VtU2VQaXNaaDFmaGg5L3hCYnANCkwrNEVCVlF2anRGV1pOczFSRzFvUFJZbHJqc2lNcWxoc1R2OENlcjlxZTFxU1N0cWMyYmxzeEZ6TXM2bFk4cC8NClByK2Jubkd6QWdNQkFBRUNnZ0VBUjRsMytqZ3kybmxseWtiSlNXQ3ZnSCs2RWtZNkRxdHd3eFlwVUpIV09sUDcNCjVtaTNWS3htY0FFT0U5V0l4S05RTnNyV0E5TnlRMFlSZjc4MnBZRGJQcEp1NHlxUjFqSTN1SVJsWlhSZU52RzcNCjNnVGpiaVBVbVRTeTBCZXY0TzFGMmZuUEdwV1ZuR2VTT1dqcnNobWExTXlocGwyV2VMRHFiSU96R2t3aHhYOXkNClRhRFd5MjErbDFpNVNGWUZTdHdXOWlhOXRORTFTTTU4WnpQWk0yK0NDdHhQVEFBQXRJRmZXUVdTbnhodUxMenMNCjNyVDRVOGNLZzJITVBXb29rOS9peWxsa0xEVXBPanhJR2tHWXdheDVnR2xvR0xZYWVoelc5Q3hobzgvc3A4WjUNCkVNNVFvczVJSTF2K21pNHhHa0RTdW4rbDYzcDN5Nm54T3pqM1h1MzRlUUtCZ1FEUDNtRWttN2lVaTlhRUxweXYNCkIxeDFlRFR2UmEwcllZMHZUaXFrYzhyUGc0NU1uOUNWRWZqdnV3YkN4M21tTExabThqZVY3ZTFHWjZJeXYreEUNCmcxeFkrUTd0RUlCb1FwWThlemg0UVYvMXRkZkhiUzNPcGdIbHVqMGd5MWxqT2QrbkxzS2RNQWRlYVF3Uy9WK2MNCk51Sks0Y3oyQWl6UXU1dHQ4WHdoOGdvU0Z3S0JnUURMNXRjZnF0VmdMQWJmMnJQbEhBLzdNcU1sWGpqNUQ0ejkNCjZmTWlCVDdOWHlYUGx6a2pJQkxOdG9OWlBCVTFzeERFb2tiNUtyTlhLTUtIaU9nTkQ0cWtDYkdnRFk2WUdaS3cNCkg4bDlLWDBaM2pwcEp0TURvQ21yQW9hSmZTUXNreGJXSDd4VlFGVzdPVWQ0dHMxZ3FDbTBUTFVxeW9lcW1EK3INCmg3WFlaa2RxeFFLQmdBK2NpZnN2M3NyNVBhRXJ4d1MyTHRGN3Q2NElzNXJBZHRRSXNOY3RBeHhXcXdkQ01XNGcNCnJXdUR4bHcya3dKUjlWa0I4LzdFb2I5WjVTcWVrMllKMzVPbkVPSHBEVnZITkhWU1k4bFVUNXFxajR3Z3ZRSDYNCkljWlpHR0l3STRSNlFqdlNIVGVrOWNpM1p2cStJTUlndFJvZW4wQVNwYjcvZUFybnlnVGFvcnI5QW9HQkFJT3QNCllOSEhqaUtjYkJnV2NjU01tZGw4T3hXL3dvVTlRSzBkYjNGUjk5dkREWFVCVU5uWk5hdDVxVnR3VExZd0hLMFANCnEwdndBbjlRQ0VoazVvN0FzYVQ3eWFUMS9GZEhkSTZmQ0l6MnhSNTJnRHcxNFdIZkJlbTFLTk1UYU5BTWNWdjQNCmhMUjlacUFRL3BIN1k2aC9FT2VwL2ZsVGI4ZUFxT1dLTDZvL2F2R05Bb0dCQUlHc0c1VExuSmlPU044SUtGU04NCmJmK3IrNkhWL2R6MkluNjhSR255MTB0OGpwbUpPbGgrdXRncGtvOXI2Y09uWGY4VHM2SFAveTBtbDl5YXhvMlANCm52c2wwcFlseFQxQy9taXJaZWxYKzFaQTltdFpHT2RxbzZhdVZUM1drcXBpb3c2WUtzbzl2Z2RHWmRWRUxiMEINCnUvdyt4UjBvN21aSEpwVEdmS09KdE53MQ0KLS0tLS1FTkQgUFJJVkFURSBLRVktLS0tLQ0K'
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
    // base64 encoded PEM cert
    auth.service.certificate = 'QmFnIEF0dHJpYnV0ZXMNCiAgICBsb2NhbEtleUlEOiBDQyBGNCBGMiBBMyBDMyBEMiAwOSBDNSAxMiBCMyA3MiA0QiBCOCA4MyBBNSA0NyA0QyAwOSAyMSBEQyANCnN1YmplY3Q9QyA9IEFVLCBTVCA9IFNvbWUtU3RhdGUsIE8gPSBJbnRlcm5ldCBXaWRnaXRzIFB0eSBMdGQNCg0KaXNzdWVyPUMgPSBBVSwgU1QgPSBTb21lLVN0YXRlLCBPID0gSW50ZXJuZXQgV2lkZ2l0cyBQdHkgTHRkDQoNCi0tLS0tQkVHSU4gQ0VSVElGSUNBVEUtLS0tLQ0KTUlJRGF6Q0NBbE9nQXdJQkFnSVVXb25VNFM0RTcxRjVZMU5zU0xYbUlhZ1dkNVl3RFFZSktvWklodmNOQVFFTA0KQlFBd1JURUxNQWtHQTFVRUJoTUNRVlV4RXpBUkJnTlZCQWdNQ2xOdmJXVXRVM1JoZEdVeElUQWZCZ05WQkFvTQ0KR0VsdWRHVnlibVYwSUZkcFpHZHBkSE1nVUhSNUlFeDBaREFlRncweE9UQTNNVEl5TVRVek1qbGFGdzB5TURBMw0KTVRFeU1UVXpNamxhTUVVeEN6QUpCZ05WQkFZVEFrRlZNUk13RVFZRFZRUUlEQXBUYjIxbExWTjBZWFJsTVNFdw0KSHdZRFZRUUtEQmhKYm5SbGNtNWxkQ0JYYVdSbmFYUnpJRkIwZVNCTWRHUXdnZ0VpTUEwR0NTcUdTSWIzRFFFQg0KQVFVQUE0SUJEd0F3Z2dFS0FvSUJBUUNsa01lQXlKbTJkMy95aEV0NHZGYjYrYjEyUGxRSDB4VGx1a1BoK2xScg0KOXJDNk5DM3dObnoySm5vbE1HclhuZVp2TlN5czFONVpSTm0yTjhQdy9QOExxeHJSenFFOFBNVC96NnN1UFhSUg0KWm5hZ2xaUklXb0NNR25pRVlDZVJHZnI4R2JpUXcwYlZEeXFuSnJaZjByS0pHbnZUNlY3QmpUdFloRWIzeXhoNA0KSmNUSnIrVDl0OEFYaldmemt6alBZdklxYmhha3FxcHd1SEVPYkh4T201cHVERTFBNVJOZm8wamcrTmZtVko5VQ0KMWR1RjVzdmE2NVQ5Q1RtdEdlbVNlUGlzWmgxZmhoOS94QmJwTCs0RUJWUXZqdEZXWk5zMVJHMW9QUllscmpzaQ0KTXFsaHNUdjhDZXI5cWUxcVNTdHFjMmJsc3hGek1zNmxZOHAvUHIrYm5uR3pBZ01CQUFHalV6QlJNQjBHQTFVZA0KRGdRV0JCU203cWFreXQwY2xxN0lnRFRWdkUrWEpaNFU5akFmQmdOVkhTTUVHREFXZ0JTbTdxYWt5dDBjbHE3SQ0KZ0RUVnZFK1hKWjRVOWpBUEJnTlZIUk1CQWY4RUJUQURBUUgvTUEwR0NTcUdTSWIzRFFFQkN3VUFBNElCQVFBYQ0KQnVqTytveU0yL0Q0SzNpS3lqVDVzbHF2UFVlVzFrZVVXYVdSVDZXRTY0VkFPbTlPZzU1bkIyOE5TSVVXampXMA0KdTJEUHF3SzJiOEFXalEveWp3S3NUMXVTdzcyQ0VEY2o3SkE1VXA5UWpBa0hIZmFoQWtOd0o5M0llcmFBdTEyVQ0KN25FRDdIN20yeGZscDVwM0dadzNHUE0rZmpBaDZLOUZIRDI0bWdGUTh4b2JPQSttVEVvV2ZIVVQrZ1pUMGxYdQ0KazFrVTJVelVOd2dwc3c4V04wNFFzWU5XcFF5d3ppUWtuZTQzNW5tdmxZOGZRc2hPSnErK0JCS0thd0xEcjk3bA0KRTBYQUxEZDZlVVhQenZ5OU1xZlozeUswRmUzMy8zbnZnUnE4QWZ3azRsbzhac2ZYWUlSTXA3b3BER0VmaUZmNQ0KM3JTTGxSZG9TNDQ4OVFZRnAyYUQNCi0tLS0tRU5EIENFUlRJRklDQVRFLS0tLS0NCkJhZyBBdHRyaWJ1dGVzDQogICAgbG9jYWxLZXlJRDogQ0MgRjQgRjIgQTMgQzMgRDIgMDkgQzUgMTIgQjMgNzIgNEIgQjggODMgQTUgNDcgNEMgMDkgMjEgREMgDQpLZXkgQXR0cmlidXRlczogPE5vIEF0dHJpYnV0ZXM+DQotLS0tLUJFR0lOIFBSSVZBVEUgS0VZLS0tLS0NCk1JSUV2Z0lCQURBTkJna3Foa2lHOXcwQkFRRUZBQVNDQktnd2dnU2tBZ0VBQW9JQkFRQ2xrTWVBeUptMmQzL3kNCmhFdDR2RmI2K2IxMlBsUUgweFRsdWtQaCtsUnI5ckM2TkMzd05uejJKbm9sTUdyWG5lWnZOU3lzMU41WlJObTINCk44UHcvUDhMcXhyUnpxRThQTVQvejZzdVBYUlJabmFnbFpSSVdvQ01HbmlFWUNlUkdmcjhHYmlRdzBiVkR5cW4NCkpyWmYwcktKR252VDZWN0JqVHRZaEViM3l4aDRKY1RKcitUOXQ4QVhqV2Z6a3pqUFl2SXFiaGFrcXFwd3VIRU8NCmJIeE9tNXB1REUxQTVSTmZvMGpnK05mbVZKOVUxZHVGNXN2YTY1VDlDVG10R2VtU2VQaXNaaDFmaGg5L3hCYnANCkwrNEVCVlF2anRGV1pOczFSRzFvUFJZbHJqc2lNcWxoc1R2OENlcjlxZTFxU1N0cWMyYmxzeEZ6TXM2bFk4cC8NClByK2Jubkd6QWdNQkFBRUNnZ0VBUjRsMytqZ3kybmxseWtiSlNXQ3ZnSCs2RWtZNkRxdHd3eFlwVUpIV09sUDcNCjVtaTNWS3htY0FFT0U5V0l4S05RTnNyV0E5TnlRMFlSZjc4MnBZRGJQcEp1NHlxUjFqSTN1SVJsWlhSZU52RzcNCjNnVGpiaVBVbVRTeTBCZXY0TzFGMmZuUEdwV1ZuR2VTT1dqcnNobWExTXlocGwyV2VMRHFiSU96R2t3aHhYOXkNClRhRFd5MjErbDFpNVNGWUZTdHdXOWlhOXRORTFTTTU4WnpQWk0yK0NDdHhQVEFBQXRJRmZXUVdTbnhodUxMenMNCjNyVDRVOGNLZzJITVBXb29rOS9peWxsa0xEVXBPanhJR2tHWXdheDVnR2xvR0xZYWVoelc5Q3hobzgvc3A4WjUNCkVNNVFvczVJSTF2K21pNHhHa0RTdW4rbDYzcDN5Nm54T3pqM1h1MzRlUUtCZ1FEUDNtRWttN2lVaTlhRUxweXYNCkIxeDFlRFR2UmEwcllZMHZUaXFrYzhyUGc0NU1uOUNWRWZqdnV3YkN4M21tTExabThqZVY3ZTFHWjZJeXYreEUNCmcxeFkrUTd0RUlCb1FwWThlemg0UVYvMXRkZkhiUzNPcGdIbHVqMGd5MWxqT2QrbkxzS2RNQWRlYVF3Uy9WK2MNCk51Sks0Y3oyQWl6UXU1dHQ4WHdoOGdvU0Z3S0JnUURMNXRjZnF0VmdMQWJmMnJQbEhBLzdNcU1sWGpqNUQ0ejkNCjZmTWlCVDdOWHlYUGx6a2pJQkxOdG9OWlBCVTFzeERFb2tiNUtyTlhLTUtIaU9nTkQ0cWtDYkdnRFk2WUdaS3cNCkg4bDlLWDBaM2pwcEp0TURvQ21yQW9hSmZTUXNreGJXSDd4VlFGVzdPVWQ0dHMxZ3FDbTBUTFVxeW9lcW1EK3INCmg3WFlaa2RxeFFLQmdBK2NpZnN2M3NyNVBhRXJ4d1MyTHRGN3Q2NElzNXJBZHRRSXNOY3RBeHhXcXdkQ01XNGcNCnJXdUR4bHcya3dKUjlWa0I4LzdFb2I5WjVTcWVrMllKMzVPbkVPSHBEVnZITkhWU1k4bFVUNXFxajR3Z3ZRSDYNCkljWlpHR0l3STRSNlFqdlNIVGVrOWNpM1p2cStJTUlndFJvZW4wQVNwYjcvZUFybnlnVGFvcnI5QW9HQkFJT3QNCllOSEhqaUtjYkJnV2NjU01tZGw4T3hXL3dvVTlRSzBkYjNGUjk5dkREWFVCVU5uWk5hdDVxVnR3VExZd0hLMFANCnEwdndBbjlRQ0VoazVvN0FzYVQ3eWFUMS9GZEhkSTZmQ0l6MnhSNTJnRHcxNFdIZkJlbTFLTk1UYU5BTWNWdjQNCmhMUjlacUFRL3BIN1k2aC9FT2VwL2ZsVGI4ZUFxT1dLTDZvL2F2R05Bb0dCQUlHc0c1VExuSmlPU044SUtGU04NCmJmK3IrNkhWL2R6MkluNjhSR255MTB0OGpwbUpPbGgrdXRncGtvOXI2Y09uWGY4VHM2SFAveTBtbDl5YXhvMlANCm52c2wwcFlseFQxQy9taXJaZWxYKzFaQTltdFpHT2RxbzZhdVZUM1drcXBpb3c2WUtzbzl2Z2RHWmRWRUxiMEINCnUvdyt4UjBvN21aSEpwVEdmS09KdE53MQ0KLS0tLS1FTkQgUFJJVkFURSBLRVktLS0tLQ0K';
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