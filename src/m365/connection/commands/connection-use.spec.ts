import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType } from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command from './connection-use.js';
import { settingsNames } from '../../../settingsNames.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';
import { ConnectionDetails } from '../../commands/ConnectionDetails.js';

describe(commands.USE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const mockContosoApplicationIdentityResponse = {
    "connectedAs": "Contoso Application",
    "connectionName": "acd6df42-10a9-4315-8928-53334f1c9d01",
    "authType": "secret",
    "appId": "39446e2e-5081-4887-980c-f285919fccca",
    "appTenant": "db308122-52f3-4241-af92-1734aa6e2e50",
    "cloudType": "Public"
  };

  const mockUserIdentityResponse = {
    "connectedAs": "alexw@contoso.com",
    "connectionName": "028de82d-7fd9-476e-a9fd-be9714280ff3",
    "authType": "deviceCode",
    "appId": "31359c7f-bd7e-475c-86db-fdb8c937548e",
    "appTenant": "common",
    "cloudType": "Public"
  };

  const connections = [
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

  before(() => {
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => settingName === settingsNames.prompt ? false : defaultValue);

    auth.connection.active = true;
    auth.connection.authType = AuthType.DeviceCode;
    auth.connection.name = '028de82d-7fd9-476e-a9fd-be9714280ff3';
    auth.connection.identityName = 'alexw@contoso.com';
    auth.connection.identityId = '028de82d-7fd9-476e-a9fd-be9714280ff3';
    auth.connection.identityTenantId = 'db308122-52f3-4241-af92-1734aa6e2e50';
    auth.connection.appId = '31359c7f-bd7e-475c-86db-fdb8c937548e';
    auth.connection.tenant = 'common';

    (auth as any)._allConnections = connections;
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

    sinon.stub(auth, 'ensureAccessToken').resolves();
  });

  afterEach(() => {
    sinonUtil.restore([
      auth.ensureAccessToken,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    auth.connection.deactivate();
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`fails with error if the connection cannot be found`, async () => {
    await assert.rejects(command.action(logger, { options: { name: 'Non-existent connection' } }),
      new CommandError(`The connection 'Non-existent connection' cannot be found.`));
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

  it(`switches to the 'Contoso Application' identity using the name option`, async () => {
    await command.action(logger, { options: { name: 'acd6df42-10a9-4315-8928-53334f1c9d01' } });
    assert(loggerLogSpy.calledOnceWithExactly(mockContosoApplicationIdentityResponse));
  });

  it(`switches to the user identity using the name option`, async () => {
    await command.action(logger, { options: { name: '028de82d-7fd9-476e-a9fd-be9714280ff3' } });
    assert(loggerLogSpy.calledOnceWithExactly(mockUserIdentityResponse));
  });

  it(`switches to the user identity using the name option (debug)`, async () => {
    await command.action(logger, { options: { name: '028de82d-7fd9-476e-a9fd-be9714280ff3', debug: true } });
    const logged = loggerLogSpy.args[0][0] as unknown as ConnectionDetails;
    assert.strictEqual(logged.connectedAs, mockUserIdentityResponse.connectedAs);
  });

  it('switches to the identity connection using prompting', async () => {
    sinon.stub(cli, 'handleMultipleResultsFound').resolves(connections[1]);

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly(mockContosoApplicationIdentityResponse));
  });

  it(`switches to the user identity using prompting`, async () => {
    sinon.stub(cli, 'handleMultipleResultsFound').resolves(connections[0]);

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly(mockUserIdentityResponse));
  });
});