import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType } from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command, { options } from './connection-set.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';

describe(commands.SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;

    sinon.stub(auth, 'ensureAccessToken').resolves();

    auth.connection.active = true;
    auth.connection.authType = AuthType.DeviceCode;
    auth.connection.name = 'Contoso';
    auth.connection.identityName = 'alexw@contoso.com';
    auth.connection.identityId = '028de82d-7fd9-476e-a9fd-be9714280ff3';
    auth.connection.identityTenantId = 'db308122-52f3-4241-af92-1734aa6e2e50';
    auth.connection.appId = '31359c7f-bd7e-475c-86db-fdb8c937548e';
    auth.connection.tenant = 'common';

    (auth as any)._allConnections = [
      {
        authType: AuthType.DeviceCode,
        active: true,
        name: 'Contoso',
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
        name: 'Fabrikam',
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

    auth.connection.name = 'Contoso';
    (auth as any)._allConnections[0].name = 'Contoso';
    (auth as any)._allConnections[1].name = 'Fabrikam';
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

  it('fails validation if newName is already an existing connection name', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ name: 'Contoso', newName: 'Fabrikam' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if name is not an existing connection name', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ name: 'NonExistent', newName: 'Contoso Application' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if name and newName are correctly set', async () => {
    const actual = await commandOptionsSchema.safeParseAsync({ name: 'Contoso', newName: 'Contoso Application' });
    assert.strictEqual(actual.success, true);
  });

  it('fails with error when restoring auth information leads to error', async () => {
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(async () => { throw 'An error has occurred'; });

    try {
      await assert.rejects(command.action(logger, { options: {} } as any),
        new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore(auth.restoreAuth);
    }
  });

  it(`Updates the 'Contoso Application' connection`, async () => {
    await command.action(logger, { options: { name: 'Fabrikam', newName: 'ContosoApplication', verbose: false, debug: false } });
    assert.strictEqual((auth as any)._allConnections[1].name, 'ContosoApplication');
  });

  it(`Updates the active user connection (debug)`, async () => {
    await command.action(logger, { options: { name: 'Contoso', newName: 'myalias', verbose: false, debug: true } });
    assert.strictEqual(auth.connection.name, 'myalias');
  });
});