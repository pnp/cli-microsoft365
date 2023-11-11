import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType, Connection } from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command from './identity-set.js';
import { Cli } from '../../../cli/Cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { settingsNames } from '../../../settingsNames.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import { CommandError } from '../../../Command.js';

describe(commands.SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const mockContosoApplicationIdentityResponse = {
    "identityName": "Contoso Application",
    "identityId": "acd6df42-10a9-4315-8928-53334f1c9d01",
    "authType": "Secret",
    "appId": "39446e2e-5081-4887-980c-f285919fccca",
    "appTenant": "db308122-52f3-4241-af92-1734aa6e2e50",
    "cloudType": "Public"
  };

  const mockUserIdentityResponse = {
    "identityName": "alexw@contoso.com",
    "identityId": "028de82d-7fd9-476e-a9fd-be9714280ff3",
    "authType": "DeviceCode",
    "appId": "31359c7f-bd7e-475c-86db-fdb8c937548e",
    "appTenant": "common",
    "cloudType": "Public"
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = Cli.getCommandInfo(command);
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
      },
      availableIdentities: [
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
          active: true,
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
      ]
    });
  });

  afterEach(() => {
    auth.service.logout();
    sinonUtil.restore([
      cli.getSettingWithDefaultValue,
      (auth as any).getConnectionInfoFromStorage,
      auth.ensureAccessToken,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid guid', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid guid', async () => {
    const actual = await command.validate({ options: { id: '0dbe7872-62f1-4b7c-b3c3-2bb71f2c63c4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if neither id or name is specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and name are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`fails with error if the identity cannot be found`, async () => {
    await assert.rejects(command.action(logger, { options: { name: 'Non-existent identity' } }), new CommandError(`The identity 'Non-existent identity' cannot be found`));
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
    await assert.doesNotReject(command.action(logger, { options: { name: 'Contoso Application' } }));
    assert(loggerLogSpy.calledOnceWithExactly(mockContosoApplicationIdentityResponse));
  });

  it(`switches to the user identity using the name option`, async () => {
    await assert.doesNotReject(command.action(logger, { options: { name: 'alexw@contoso.com' } }));
    assert(loggerLogSpy.calledOnceWithExactly(mockUserIdentityResponse));
  });

  it(`switches to the 'Contoso Application' identity using the id option`, async () => {
    await assert.doesNotReject(command.action(logger, { options: { id: 'acd6df42-10a9-4315-8928-53334f1c9d01' } }));
    assert(loggerLogSpy.calledOnceWithExactly(mockContosoApplicationIdentityResponse));
  });

  it(`switches to the user identity using the id option`, async () => {
    await assert.doesNotReject(command.action(logger, { options: { id: '028de82d-7fd9-476e-a9fd-be9714280ff3' } }));
    assert(loggerLogSpy.calledOnceWithExactly(mockUserIdentityResponse));
  });

  it(`switches to the user identity using the name option (debug)`, async () => {
    await assert.doesNotReject(command.action(logger, { options: { name: 'alexw@contoso.com', debug: true } }));
    const logged = loggerLogSpy.args[0][0] as unknown as Connection;
    assert(loggerLogSpy.calledOnce);
    assert.strictEqual(logged.identityName, mockUserIdentityResponse.identityName);
    assert.strictEqual(logged.identityId, mockUserIdentityResponse.identityId);
  });

  it(`fails refreshing access token while switching (debug)`, async () => {
    const mockError = new Error('MockErrorMessage');
    sinonUtil.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').rejects(mockError);

    await assert.rejects(command.action(logger, { options: { id: '028de82d-7fd9-476e-a9fd-be9714280ff3', debug: true } }), new CommandError('Your login has expired. Sign in again to continue. MockErrorMessage'));
  });

  it('handles selecting single result when multiple identities with the specified name found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
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
      },
      availableIdentities: [
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
              expiresOn: (new Date()).toISOString(),
              accessToken: 'abc'
            }
          }
        },
        {
          authType: AuthType.Secret,
          active: true,
          identityName: 'Contoso Application',
          identityId: '46657f7d-a133-43f1-8721-6e4f53b43c97',
          appId: '0445b0a6-88ff-499b-b91b-c181d0c24772',
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
      ]
    });

    await assert.rejects(command.action(logger, { options: { name: 'Contoso Application' } }), new CommandError(`Multiple identities with 'Contoso Application' found. Found: acd6df42-10a9-4315-8928-53334f1c9d01, 46657f7d-a133-43f1-8721-6e4f53b43c97.`));
  });

  it('handles selecting single result when multiple identities with the specified name found and cli is set to prompt', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return true;
      }

      return defaultValue;
    });

    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
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
      },
      availableIdentities: [
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
              expiresOn: (new Date()).toISOString(),
              accessToken: 'abc'
            }
          }
        },
        {
          authType: AuthType.Secret,
          active: true,
          identityName: 'Contoso Application',
          identityId: '46657f7d-a133-43f1-8721-6e4f53b43c97',
          appId: '0445b0a6-88ff-499b-b91b-c181d0c24772',
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
      ]
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves({
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
          expiresOn: (new Date()).toISOString(),
          accessToken: 'abc'
        }
      }
    });

    await assert.doesNotReject(command.action(logger, { options: { name: 'Contoso Application' } }));
    assert(loggerLogSpy.calledOnceWithExactly(mockContosoApplicationIdentityResponse));
  });
});