import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType } from '../../Auth.js';
import { CommandError } from '../../Command.js';
import { Logger } from '../../cli/Logger.js';
import { telemetry } from '../../telemetry.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import commands from './commands.js';
import command from './status.js';

describe(commands.STATUS, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  const mockUserIdentityResponse = {
    "connectedAs": "alexw@contoso.com",
    "identityName": "alexw@contoso.com",
    "identityId": "028de82d-7fd9-476e-a9fd-be9714280ff3",
    "authType": "DeviceCode",
    "appId": "31359c7f-bd7e-475c-86db-fdb8c937548e",
    "appTenant": "common",
    "cloudType": "Public"
  };

  before(() => {
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').resolves('');
    sinon.stub(session, 'getId').resolves('');
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(auth as any, 'getConnectionInfoFromStorage').resolves({
      authType: AuthType.DeviceCode,
      active: true,
      identityName: 'alexw@contoso.com',
      identityId: '028de82d-7fd9-476e-a9fd-be9714280ff3',
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      certificateType: CertificateType.Unknown,
      accessTokens: {
        'https://graph.microsoft.com': {
          expiresOn: (new Date()).toISOString(),
          accessToken: 'abc'
        }
      }
    });
    sinon.stub(auth as any, 'getAllConnectionsFromStorage').resolves([
      {
        authType: AuthType.DeviceCode,
        active: true,
        identityName: 'alexw@contoso.com',
        identityId: '028de82d-7fd9-476e-a9fd-be9714280ff3',
        appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        tenant: 'common',
        cloudType: CloudType.Public,
        certificateType: CertificateType.Unknown,
        accessTokens: {
          'https://graph.microsoft.com': {
            expiresOn: (new Date()).toISOString(),
            accessToken: 'abc'
          }
        }
      },
      {
        authType: AuthType.Secret,
        active: false,
        identityName: 'Contoso Application',
        identityId: 'acd6df42-10a9-4315-8928-53334f1c9d01',
        appId: '39446e2e-5081-4887-980c-f285919fccca',
        tenant: 'db308122-52f3-4241-af92-1734aa6e2e50',
        cloudType: CloudType.Public,
        certificateType: CertificateType.Unknown,
        accessTokens: {
          'https://graph.microsoft.com': {
            expiresOn: (new Date()).toISOString(),
            accessToken: 'abc'
          }
        }
      }
    ]);
  });

  afterEach(() => {
    auth.service.logout();
    sinonUtil.restore([
      auth.ensureAccessToken,
      (auth as any).getConnectionInfoFromStorage,
      (auth as any).getAllConnectionsFromStorage
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.STATUS), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows logged out status when not logged in', async () => {
    auth.service.logout();
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').resolves();
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith('Logged out'));
    sinonUtil.restore(auth.restoreAuth);
  });

  it('shows logged out status when not logged in (verbose)', async () => {
    auth.service.logout();
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').resolves();
    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogToStderrSpy.calledWith('Logged out from Microsoft 365'));
    sinonUtil.restore(auth.restoreAuth);
  });

  it('shows logged out status when not logged in, but identities available', async () => {
    sinon.stub(auth, 'ensureAccessToken').resolves();
    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
    sinonUtil.restore((auth as any).getAllConnectionsFromStorage);
    sinon.stub(auth as any, 'getConnectionInfoFromStorage').resolves({
      authType: AuthType.DeviceCode,
      active: false,
      identityName: undefined,
      identityId: undefined,
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      certificateType: CertificateType.Unknown,
      accessTokens: {}
    });
    sinon.stub(auth as any, 'getAllConnectionsFromStorage').resolves([
      {
        authType: AuthType.Secret,
        active: true,
        identityName: 'Contoso Application',
        identityId: 'acd6df42-10a9-4315-8928-53334f1c9d01',
        appId: '39446e2e-5081-4887-980c-f285919fccca',
        tenant: 'db308122-52f3-4241-af92-1734aa6e2e50',
        cloudType: CloudType.Public,
        certificateType: CertificateType.Unknown,
        accessTokens: {
          'https://graph.microsoft.com': {
            expiresOn: '123',
            accessToken: 'abc'
          }
        }
      }
    ]);
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith('Logged out, signed in identities available'));
  });

  it('shows logged out status when not logged in, but identities available (verbose)', async () => {
    sinon.stub(auth, 'ensureAccessToken').resolves();
    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
    sinonUtil.restore((auth as any).getAllConnectionsFromStorage);
    sinon.stub(auth as any, 'getConnectionInfoFromStorage').resolves({
      authType: AuthType.DeviceCode,
      active: false,
      identityName: undefined,
      identityId: undefined,
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      certificateType: CertificateType.Unknown,
      accessTokens: {}
    });
    sinon.stub(auth as any, 'getAllConnectionsFromStorage').resolves([
      {
        authType: AuthType.Secret,
        active: true,
        identityName: 'Contoso Application',
        identityId: 'acd6df42-10a9-4315-8928-53334f1c9d01',
        appId: '39446e2e-5081-4887-980c-f285919fccca',
        tenant: 'db308122-52f3-4241-af92-1734aa6e2e50',
        cloudType: CloudType.Public,
        certificateType: CertificateType.Unknown,
        accessTokens: {
          'https://graph.microsoft.com': {
            expiresOn: '123',
            accessToken: 'abc'
          }
        }
      }
    ]);

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogToStderrSpy.calledWith('Logged out from Microsoft 365, signed in identities available'));
  });

  it('shows logged out status when the refresh token is expired', async () => {
    auth.service.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: 'abc',
      accessToken: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ing0NTh4eU9wbHNNMkg3TlhrMlN4MTd4MXVwYyIsImtpZCI6Ing0NTh4eU9wbHNNMkg3TlhrN1N4MTd4MXVwYyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLndpbmRvd3MubmV0IiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvY2FlZTMyZTYtNDA1ZC00MjRhLTljZjEtMjA3MWQwNDdmMjk4LyIsImlhdCI6MTUxNTAwNDc4NCwibmJmIjoxNTE1MDA0Nzg0LCJleHAiOjE1MTUwMDg2ODQsImFjciI6IjEiLCJhaW8iOiJBQVdIMi84R0FBQUFPN3c0TDBXaHZLZ1kvTXAxTGJMWFdhd2NpOEpXUUpITmpKUGNiT2RBM1BvPSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiIwNGIwNzc5NS04ZGRiLTQ2MWEtYmJlZS0wMmY5ZTFiZjdiNDYiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2huIiwiaXBhZGRyIjoiOC44LjguOCIsIm5hbWUiOiJKb2huIERvZSIsIm9pZCI6ImYzZTU5NDkxLWZjMWEtNDdjYy1hMWYwLTk1ZWQ0NTk4MzcxNyIsInB1aWQiOiIxMDk0N0ZGRUE2OEJDQ0NFIiwic2NwIjoiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwic3ViIjoiemZicmtUV1VQdEdWUUg1aGZRckpvVGp3TTBrUDRsY3NnLTJqeUFJb0JuOCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImNhZWUzM2U2LTQwNWQtNDU0YS05Y2YxLTMwNzFkMjQxYTI5OCIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AY29udG9zby5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBjb250b3NvLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImFUZVdpelVmUTBheFBLMVRUVXhsQUEiLCJ2ZXIiOiIxLjAifQ==.abc'
    };

    sinon.stub(auth, 'ensureAccessToken').rejects(new Error('Error'));
    await assert.rejects(command.action(logger, { options: {} }), new CommandError(`Your login has expired. Sign in again to continue. Error`));
  });

  it('shows logged out status when refresh token is expired (debug)', async () => {
    auth.service.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: 'abc',
      accessToken: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ing0NTh4eU9wbHNNMkg3TlhrMlN4MTd4MXVwYyIsImtpZCI6Ing0NTh4eU9wbHNNMkg3TlhrN1N4MTd4MXVwYyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLndpbmRvd3MubmV0IiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvY2FlZTMyZTYtNDA1ZC00MjRhLTljZjEtMjA3MWQwNDdmMjk4LyIsImlhdCI6MTUxNTAwNDc4NCwibmJmIjoxNTE1MDA0Nzg0LCJleHAiOjE1MTUwMDg2ODQsImFjciI6IjEiLCJhaW8iOiJBQVdIMi84R0FBQUFPN3c0TDBXaHZLZ1kvTXAxTGJMWFdhd2NpOEpXUUpITmpKUGNiT2RBM1BvPSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiIwNGIwNzc5NS04ZGRiLTQ2MWEtYmJlZS0wMmY5ZTFiZjdiNDYiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2huIiwiaXBhZGRyIjoiOC44LjguOCIsIm5hbWUiOiJKb2huIERvZSIsIm9pZCI6ImYzZTU5NDkxLWZjMWEtNDdjYy1hMWYwLTk1ZWQ0NTk4MzcxNyIsInB1aWQiOiIxMDk0N0ZGRUE2OEJDQ0NFIiwic2NwIjoiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwic3ViIjoiemZicmtUV1VQdEdWUUg1aGZRckpvVGp3TTBrUDRsY3NnLTJqeUFJb0JuOCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImNhZWUzM2U2LTQwNWQtNDU0YS05Y2YxLTMwNzFkMjQxYTI5OCIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AY29udG9zby5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBjb250b3NvLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImFUZVdpelVmUTBheFBLMVRUVXhsQUEiLCJ2ZXIiOiIxLjAifQ==.abc'
    };

    const error = new Error('Error');
    sinon.stub(auth, 'ensureAccessToken').rejects(error);
    await assert.rejects(command.action(logger, { options: { debug: true } }), new CommandError(`Your login has expired. Sign in again to continue. Error`));
    assert(loggerLogToStderrSpy.calledWith(error));
  });

  it('shows logged in status when logged in', async () => {
    sinon.stub(auth, 'ensureAccessToken').resolves();
    await assert.doesNotReject(command.action(logger, { options: {} }));
    assert(loggerLogSpy.calledWith(mockUserIdentityResponse));
  });

  it('shows logged in status when logged in (debug)', async () => {
    sinon.stub(auth, 'ensureAccessToken').resolves();
    await assert.doesNotReject(command.action(logger, { options: { debug: true } }));
    assert(loggerLogToStderrSpy.calledWith({
      ...mockUserIdentityResponse,
      accessTokens: '{\n  "https://graph.microsoft.com": {\n    "expiresOn": "123",\n    "accessToken": "abc"\n  }\n}'
    }));
  });

  it('correctly handles error when restoring auth', async () => {
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
    sinonUtil.restore(auth.restoreAuth);
  });
});
