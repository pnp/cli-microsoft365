import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType } from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import { telemetry } from '../../telemetry.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import commands from './commands.js';
import command from './logout.js';
import { Cli } from '../../cli/Cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { settingsNames } from '../../settingsNames.js';

describe(commands.LOGOUT, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let authClearConnectionInfoStub: sinon.SinonStub;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    authClearConnectionInfoStub = sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      (auth as any).getConnectionInfoFromStorage,
      (auth as any).getAllConnectionsFromStorage,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.logout();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LOGOUT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the identityId is not a valid guid', async () => {
    const actual = await command.validate({ options: { identityId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the identityId is a valid guid', async () => {
    const actual = await command.validate({ options: { identityId: '0dbe7872-62f1-4b7c-b3c3-2bb71f2c63c4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if neither identityId or identityName is specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if identityId and identityName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });
    const actual = await command.validate({ options: { identityId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', identityName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('logs out from Microsoft 365 when logged in', async () => {
    await command.action(logger, { options: { debug: true } });
    assert(!auth.service.active);
  });

  it('logs out from Microsoft 365 when not logged in', async () => {
    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
    await command.action(logger, { options: { debug: true } });
    assert(!auth.service.active);
  });

  it('logs out from specific user account by name when logged in', async () => {
    await assert.doesNotReject(command.action(logger, { options: { debug: true, identityName: 'alexw@contoso.com' } }));
    assert(!auth.service.active);
  });

  it('logs out from specific user account by id when logged in', async () => {
    await assert.doesNotReject(command.action(logger, { options: { debug: true, identityId: '028de82d-7fd9-476e-a9fd-be9714280ff3' } }));
    assert(!auth.service.active);
  });

  it('clears persisted connection info when logging out', async () => {
    await assert.doesNotReject(command.action(logger, { options: { debug: true } }));
    assert(authClearConnectionInfoStub.called);
  });

  it('correctly handles error while clearing persisted connection info', async () => {
    sinonUtil.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    const logoutSpy = sinon.spy(auth.service, 'logout');

    try {
      await command.action(logger, { options: {} });
      assert(logoutSpy.called);
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo,
        auth.service.logout
      ]);
    }
  });

  it('correctly handles error while clearing persisted connection info (debug)', async () => {
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    const logoutSpy = sinon.spy(auth.service, 'logout');

    try {
      await command.action(logger, { options: { debug: true } });
      assert(logoutSpy.called);
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo,
        auth.service.logout
      ]);
    }
  });

  it('correctly handles error when restoring auth information', async () => {
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));

    try {
      await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        auth.restoreAuth,
        auth.clearConnectionInfo
      ]);
    }
  });

  it('correctly handles error when not finding an identity by name in the list of available identities', async () => {
    await assert.rejects(command.action(logger, { options: { identityName: 'non-existent identity' } } as any), new CommandError(`The identity 'non-existent identity' cannot be found.`));
  });

  it('correctly handles error when not finding an identity by id in the list of available identities', async () => {
    await assert.rejects(command.action(logger, { options: { identityId: 'ecd7e376-31cb-4b7e-b59c-0272195fda80' } } as any), new CommandError(`The identity 'ecd7e376-31cb-4b7e-b59c-0272195fda80' cannot be found.`));
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

    await assert.rejects(command.action(logger, { options: { identityName: 'Contoso Application' } }), new CommandError(`Multiple identities with 'Contoso Application' found. Found: acd6df42-10a9-4315-8928-53334f1c9d01, 46657f7d-a133-43f1-8721-6e4f53b43c97.`));
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

    await assert.doesNotReject(command.action(logger, { options: { identityName: 'Contoso Application' } }));
    assert(!auth.service.active);
  });
});
