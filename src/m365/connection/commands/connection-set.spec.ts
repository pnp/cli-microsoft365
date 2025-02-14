import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType } from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command from './connection-set.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';

describe(commands.SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);

    sinon.stub(auth, 'ensureAccessToken').resolves();

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
            expiresOn: '2024-01-22T20:22:54.814Z',
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

    auth.connection.name = '028de82d-7fd9-476e-a9fd-be9714280ff3';
    (auth as any)._allConnections[0].name = '028de82d-7fd9-476e-a9fd-be9714280ff3';
    (auth as any)._allConnections[1].name = 'acd6df42-10a9-4315-8928-53334f1c9d01';
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
    assert.strictEqual(command.name, commands.SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if newName is the same as name', async () => {
    const actual = await command.validate({ options: { name: 'test', newName: 'test' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if name and newName are correctly set', async () => {
    const actual = await command.validate({ options: { name: 'oldname', newName: 'newname' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it(`fails with error if the connection cannot be found`, async () => {
    await assert.rejects(command.action(logger, { options: { name: 'Non-existent connection', newName: 'something new' } }), new CommandError(`The connection 'Non-existent connection' cannot be found.`));
  });

  it(`fails with error if the newName is already in use`, async () => {
    await assert.rejects(command.action(logger, { options: { name: 'acd6df42-10a9-4315-8928-53334f1c9d01', newName: '028de82d-7fd9-476e-a9fd-be9714280ff3' } }), new CommandError(`The connection name '028de82d-7fd9-476e-a9fd-be9714280ff3' is already in use`));
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

  it(`Updates the 'Contoso Application' connection`, async () => {
    await command.action(logger, { options: { name: 'acd6df42-10a9-4315-8928-53334f1c9d01', newName: 'ContosoApplication' } });
    assert.strictEqual((auth as any)._allConnections[1].name, 'ContosoApplication');
  });

  it(`Updates the active user connection (debug)`, async () => {
    await command.action(logger, { options: { name: '028de82d-7fd9-476e-a9fd-be9714280ff3', newName: 'myalias', debug: true } });
    assert.strictEqual(auth.connection.name, 'myalias');
  });
});