import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType } from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { spo } from '../../../utils/spo.js';
import commands from '../commands.js';
import command from './connection-list.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import { CommandError } from '../../../Command.js';

describe(commands.LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const mockListResponse = [
    {
      "active": true,
      "authType": "deviceCode",
      "connectedAs": "alexw@contoso.com",
      "name": "028de82d-7fd9-476e-a9fd-be9714280ff3"
    },
    {
      "active": false,
      "authType": "secret",
      "connectedAs": "Contoso Application",
      "name": "acd6df42-10a9-4315-8928-53334f1c9d01"
    }
  ];

  before(() => {
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    auth.connection.active = true;
    auth.connection.authType = AuthType.DeviceCode;
    auth.connection.name = '028de82d-7fd9-476e-a9fd-be9714280ff3';
    auth.connection.identityName = 'alexw@contoso.com';
    auth.connection.identityId = '028de82d-7fd9-476e-a9fd-be9714280ff3';
    auth.connection.identityTenantId = 'db308122-52f3-4241-af92-1734aa6e2e50';
    auth.connection.appId = '31359c7f-bd7e-475c-86db-fdb8c937548e';
    auth.connection.tenant = 'common';

    (auth as any)._allConnections = [
      {
        authType: AuthType.DeviceCode,
        active: true,
        name: '028de82d-7fd9-476e-a9fd-be9714280ff3',
        identityName: 'alexw@contoso.com',
        identityId: '028de82d-7fd9-476e-a9fd-be9714280ff3',
        identityTenantId: 'db308122-52f3-4241-af92-1734aa6e2e50',
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
        active: true,
        name: 'acd6df42-10a9-4315-8928-53334f1c9d01',
        identityName: 'Contoso Application',
        identityId: 'acd6df42-10a9-4315-8928-53334f1c9d01',
        identityTenantId: 'db308122-52f3-4241-af92-1734aa6e2e50',
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
    ];
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
  });

  afterEach(() => {
  });

  after(() => {
    sinon.restore();
    auth.connection.deactivate();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'connectedAs', 'authType', 'active']);
  });

  it('shows a list of connections', async () => {
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly(mockListResponse));
  });

  it('shows an empty list of connections', async () => {
    sinon.stub(auth, 'getAllConnections').resolves([]);
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly([]));
    sinonUtil.restore((auth as any).getAllConnections);
  });

  it('shows an list with one connection', async () => {
    const mockConnectionsList = [
      {
        authType: AuthType.DeviceCode,
        active: true,
        name: '028de82d-7fd9-476e-a9fd-be9714280ff3',
        identityName: 'alexw@contoso.com',
        identityId: '028de82d-7fd9-476e-a9fd-be9714280ff3',
        identityTenantId: 'db308122-52f3-4241-af92-1734aa6e2e50',
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
      }
    ];

    const mockConnectionsResponse = [
      {
        "active": true,
        "authType": "deviceCode",
        "connectedAs": "alexw@contoso.com",
        "name": "028de82d-7fd9-476e-a9fd-be9714280ff3"
      }
    ];

    sinon.stub((auth as any), 'getAllConnections').resolves(mockConnectionsList);
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly(mockConnectionsResponse));
    sinonUtil.restore((auth as any).getAllConnections);
  });

  it('shows a list of connections (debug)', async () => {
    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledOnceWithExactly(mockListResponse));
  });

  it('fails with error when restoring auth information leads to error', async () => {
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));

    try {
      await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore(auth.restoreAuth);
    }
  });
});