import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType } from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command from './connection-remove.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { settingsNames } from '../../../settingsNames.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';

describe(commands.REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);


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

    sinon.stub(auth, 'ensureAccessToken').resolves();
  });

  afterEach(() => {
    sinonUtil.restore([
      auth.ensureAccessToken,
      auth.removeConnectionInfo,
      cli.getSettingWithDefaultValue,
      cli.promptForConfirmation,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    auth.connection.deactivate();
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if name is not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`fails with error if the connection cannot be found`, async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await assert.rejects(command.action(logger, { options: { name: 'Non-existent connection' } }), new CommandError(`The connection 'Non-existent connection' cannot be found.`));
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

  it(`removes the 'Contoso Application' connection when prompt is confirmed`, async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    const removeStub = sinon.stub(auth, 'removeConnectionInfo').resolves();
    await command.action(logger, { options: { name: 'acd6df42-10a9-4315-8928-53334f1c9d01' } });
    assert(removeStub.calledOnce);
  });

  it(`removes the 'Contoso Application' connection and not prompting for confirmation`, async () => {
    const removeStub = sinon.stub(auth, 'removeConnectionInfo').resolves();
    await command.action(logger, { options: { name: 'acd6df42-10a9-4315-8928-53334f1c9d01', force: true } });
    assert(removeStub.calledOnce);
  });


  it('aborts removing the connection when prompt not confirmed', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    const removeStub = sinon.stub(auth, 'removeConnectionInfo').resolves();

    await command.action(logger, {
      options: {
        name: 'acd6df42-10a9-4315-8928-53334f1c9d01'
      }
    });
    assert(removeStub.notCalled);
  });

  it(`removes the user connection when prompt is confirmed (debug)`, async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    const removeStub = sinon.stub(auth, 'removeConnectionInfo').resolves();

    await command.action(logger, { options: { name: '028de82d-7fd9-476e-a9fd-be9714280ff3', debug: true } });
    assert(removeStub.calledOnce);
  });
});