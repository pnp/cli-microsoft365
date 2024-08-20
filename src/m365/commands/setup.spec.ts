import assert from 'assert';
import Configstore from 'configstore';
import sinon from 'sinon';
import auth from '../../Auth.js';
import { cli } from '../../cli/cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { settingsNames } from '../../settingsNames.js';
import { telemetry } from '../../telemetry.js';
import { accessToken } from '../../utils/accessToken.js';
import { entraApp } from '../../utils/entraApp.js';
import { CheckStatus, formatting } from '../../utils/formatting.js';
import { pid } from '../../utils/pid.js';
import { ConfirmationConfig, SelectionConfig } from '../../utils/prompt.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import commands from './commands.js';
import command, { CliExperience, CliUsageMode, EntraAppConfig, HelpMode, NewEntraAppScopes, Preferences, SettingNames } from './setup.js';
import { interactivePreset, powerShellPreset, scriptingPreset } from './setupPresets.js';

describe(commands.SETUP, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let config: Configstore;
  let configSetSpy: sinon.SinonSpy;
  let configDeleteSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = cli.getCommandInfo(command);
    config = cli.getConfig();
    configDeleteSpy = sinon.stub(config, 'delete').callsFake(() => { });
    configSetSpy = sinon.stub(config, 'set').callsFake(() => { });
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).answers = {};
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).configureSettings,
      cli.promptForConfirmation,
      cli.promptForInput,
      cli.promptForSelection,
      pid.isPowerShell,
      auth.clearConnectionInfo,
      auth.ensureAccessToken,
      accessToken.getTenantIdFromAccessToken,
      entraApp.resolveApis,
      entraApp.createAppRegistration,
      entraApp.grantAdminConsent
    ]);
    configSetSpy.resetHistory();
    configDeleteSpy.resetHistory();
    auth.connection.accessTokens = {};
  });

  after(() => {
    sinonUtil.restore([
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId,
      config.delete,
      config.set
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SETUP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets correct settings for interactive, beginner', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Interactively;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Beginner;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });

    await command.action(logger, { options: {} });

    assert(configSetSpy.calledWith(settingsNames.helpMode, HelpMode.Full), 'Incorrect help mode');
    Object.keys(interactivePreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (interactivePreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('sets correct settings for interactive, proficient', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Interactively;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });

    await command.action(logger, { options: {} });

    assert(configSetSpy.calledWith(settingsNames.helpMode, HelpMode.Options), 'Incorrect help mode');
    Object.keys(interactivePreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (interactivePreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('sets correct settings for scripting, non-PowerShell, beginner', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Beginner;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return false;
        default: //summary
          return true;
      }
    });

    await command.action(logger, { options: {} });

    assert(configSetSpy.calledWith(settingsNames.helpMode, HelpMode.Full), 'Incorrect help mode');
    Object.keys(scriptingPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (scriptingPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('sets correct settings for scripting, PowerShell, beginner', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Beginner;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });

    await command.action(logger, { options: {} });

    assert(configSetSpy.calledWith(settingsNames.helpMode, HelpMode.Full), 'Incorrect help mode');
    Object.keys(scriptingPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (scriptingPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
    Object.keys(powerShellPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (powerShellPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('sets correct settings for scripting, non-PowerShell, proficient', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return false;
        default: //summary
          return true;
      }
    });

    await command.action(logger, { options: {} });

    assert(configSetSpy.calledWith(settingsNames.helpMode, HelpMode.Options), 'Incorrect help mode');
    Object.keys(scriptingPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (scriptingPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('sets correct settings for scripting, PowerShell, proficient', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });

    await command.action(logger, { options: {} });

    assert(configSetSpy.calledWith(settingsNames.helpMode, HelpMode.Options), 'Incorrect help mode');
    Object.keys(scriptingPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (scriptingPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
    Object.keys(powerShellPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (powerShellPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it(`doesn't apply settings when not confirmed`, async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Beginner;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return false;
        default: //summary
          return false;
      }
    });
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.notCalled);
  });

  it('sets correct settings for interactive, non-PowerShell via option', async () => {
    sinon.stub(pid, 'isPowerShell').returns(false);

    await command.action(logger, { options: { interactive: true } });

    Object.keys(interactivePreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (interactivePreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('sets correct settings for scripting, non-PowerShell via option', async () => {
    sinon.stub(pid, 'isPowerShell').returns(false);

    await command.action(logger, { options: { scripting: true } });

    Object.keys(scriptingPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (scriptingPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('sets correct settings for interactive, PowerShell via option', async () => {
    sinon.stub(pid, 'isPowerShell').returns(true);

    await command.action(logger, { options: { interactive: true } });

    Object.keys(interactivePreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (interactivePreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
    Object.keys(powerShellPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (powerShellPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('sets correct settings for scripting, PowerShell via option', async () => {
    sinon.stub(pid, 'isPowerShell').returns(true);

    await command.action(logger, { options: { scripting: true } });

    Object.keys(scriptingPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (scriptingPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
    Object.keys(powerShellPreset).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (powerShellPreset as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('skips configuring Entra app when specified via args', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });

    await command.action(logger, { options: { skipApp: true } });

    const expected: SettingNames = {
      clientId: '00000000-0000-0000-0000-000000000000',
      tenantId: '00000000-0000-0000-0000-000000000000',
      clientSecret: '',
      clientCertificateFile: '',
      clientCertificateBase64Encoded: ''
    };
    Object.keys(expected).forEach(setting => {
      assert(!configSetSpy.calledWith(setting), `Modified ${setting}`);
    });
  });

  it('configures existing public Entra app', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        case 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?':
          return EntraAppConfig.UseExisting;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });
    sinon.stub(cli, 'promptForInput').callsFake(async (config: { message: string }): Promise<string> => {
      switch (config.message) {
        case 'Client ID:':
          return '00000000-0000-0000-0000-000000000000';
        case 'Tenant ID (leave common if the app is multitenant):':
          return '00000000-0000-0000-0000-000000000000';
        default:
          return '';
      }
    });

    await command.action(logger, { options: {} });

    const expected: SettingNames = {
      clientId: '00000000-0000-0000-0000-000000000000',
      tenantId: '00000000-0000-0000-0000-000000000000',
      clientSecret: '',
      clientCertificateFile: '',
      clientCertificateBase64Encoded: ''
    };
    Object.keys(expected).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (expected as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('configures existing Entra app with secret', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        case 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?':
          return EntraAppConfig.UseExisting;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });
    sinon.stub(cli, 'promptForInput').callsFake(async (config: { message: string }): Promise<string> => {
      switch (config.message) {
        case 'Client ID:':
          return '00000000-0000-0000-0000-000000000000';
        case 'Tenant ID (leave common if the app is multitenant):':
          return '00000000-0000-0000-0000-000000000000';
        case 'Client secret (leave empty if you use a certificate or a public client):':
          return 'secret';
        default:
          return '';
      }
    });

    await command.action(logger, { options: {} });

    const expected: SettingNames = {
      clientId: '00000000-0000-0000-0000-000000000000',
      tenantId: '00000000-0000-0000-0000-000000000000',
      clientSecret: 'secret'
    };
    Object.keys(expected).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (expected as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('configures existing Entra app with base64 cert', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        case 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?':
          return EntraAppConfig.UseExisting;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });
    sinon.stub(cli, 'promptForInput').callsFake(async (config: { message: string }): Promise<string> => {
      switch (config.message) {
        case 'Client ID:':
          return '00000000-0000-0000-0000-000000000000';
        case 'Tenant ID (leave common if the app is multitenant):':
          return '00000000-0000-0000-0000-000000000000';
        case 'Base64-encoded certificate string:':
          return 'base64';
        default:
          return '';
      }
    });

    await command.action(logger, { options: {} });

    const expected: SettingNames = {
      clientId: '00000000-0000-0000-0000-000000000000',
      tenantId: '00000000-0000-0000-0000-000000000000',
      clientSecret: '',
      clientCertificateFile: '',
      clientCertificateBase64Encoded: 'base64'
    };
    Object.keys(expected).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (expected as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('configures existing Entra app with file cert', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        case 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?':
          return EntraAppConfig.UseExisting;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });
    sinon.stub(cli, 'promptForInput').callsFake(async (config: { message: string }): Promise<string> => {
      switch (config.message) {
        case 'Client ID:':
          return '00000000-0000-0000-0000-000000000000';
        case 'Tenant ID (leave common if the app is multitenant):':
          return '00000000-0000-0000-0000-000000000000';
        case 'Path to the client certificate file (leave empty if you want to specify a base64-encoded certificate string):':
          return 'file';
        default:
          return '';
      }
    });

    await command.action(logger, { options: {} });

    const expected: SettingNames = {
      clientId: '00000000-0000-0000-0000-000000000000',
      tenantId: '00000000-0000-0000-0000-000000000000',
      clientSecret: '',
      clientCertificateFile: 'file'
    };
    Object.keys(expected).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (expected as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('configures existing Entra app with file cert secured with password', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        case 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?':
          return EntraAppConfig.UseExisting;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });
    sinon.stub(cli, 'promptForInput').callsFake(async (config: { message: string }): Promise<string> => {
      switch (config.message) {
        case 'Client ID:':
          return '00000000-0000-0000-0000-000000000000';
        case 'Tenant ID (leave common if the app is multitenant):':
          return '00000000-0000-0000-0000-000000000000';
        case 'Path to the client certificate file (leave empty if you want to specify a base64-encoded certificate string):':
          return 'file';
        case 'Password for the client certificate (leave empty if the certificate is not password-protected):':
          return 'password';
        default:
          return '';
      }
    });

    await command.action(logger, { options: {} });

    const expected: SettingNames = {
      clientId: '00000000-0000-0000-0000-000000000000',
      tenantId: '00000000-0000-0000-0000-000000000000',
      clientSecret: '',
      clientCertificateFile: 'file',
      clientCertificatePassword: 'password'
    };
    Object.keys(expected).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (expected as any)[setting]), `Incorrect setting for ${setting}`);
    });
  });

  it('creates a new Entra app with minimal scopes', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        case 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?':
          return EntraAppConfig.Create;
        case 'What scopes should the new app registration have?':
          return NewEntraAppScopes.Minimal;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'ensureAccessToken').resolves();
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').returns('00000000-0000-0000-0000-000000000003');
    const scopes = [
      {
        resourceAppId: '00000000-0000-0000-0000-000000000000',
        resourceAccess: [
          {
            id: '00000000-0000-0000-0000-000000000000',
            type: 'Minimal'
          }
        ]
      }
    ];
    sinon.stub(entraApp, 'resolveApis').resolves(scopes);
    const createAppRegistrationSpy = sinon.stub(entraApp, 'createAppRegistration').resolves({
      appId: '00000000-0000-0000-0000-000000000001',
      id: '00000000-0000-0000-0000-000000000002',
      tenantId: '00000000-0000-0000-0000-000000000003',
      requiredResourceAccess: scopes
    });
    sinon.stub(entraApp, 'grantAdminConsent').resolves();
    auth.connection.accessTokens[auth.defaultResource] = {
      accessToken: 'abc',
      expiresOn: new Date().toString()
    };

    await command.action(logger, { options: {} });

    const expected: SettingNames = {
      clientId: '00000000-0000-0000-0000-000000000001',
      tenantId: '00000000-0000-0000-0000-000000000003'
    };
    const deleted: SettingNames = {
      clientSecret: '',
      clientCertificateFile: '',
      clientCertificateBase64Encoded: '',
      clientCertificatePassword: ''
    };
    assert.deepEqual(createAppRegistrationSpy.getCall(0).args[0].apis, scopes);
    Object.keys(expected).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (expected as any)[setting]), `Incorrect setting for ${setting}`);
    });
    Object.keys(deleted).forEach(setting => {
      assert(configDeleteSpy.calledWith(setting), `Not deleted ${setting}`);
    });
  });

  it('creates a new Entra app with all scopes (verbose)', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        case 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?':
          return EntraAppConfig.Create;
        case 'What scopes should the new app registration have?':
          return NewEntraAppScopes.All;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        default: //summary
          return true;
      }
    });
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'ensureAccessToken').resolves();
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').returns('00000000-0000-0000-0000-000000000003');
    const scopes = [
      {
        resourceAppId: '00000000-0000-0000-0000-000000000000',
        resourceAccess: [
          {
            id: '00000000-0000-0000-0000-000000000000',
            type: 'All'
          }
        ]
      }
    ];
    sinon.stub(entraApp, 'resolveApis').resolves(scopes);
    const createAppRegistrationSpy = sinon.stub(entraApp, 'createAppRegistration').resolves({
      appId: '00000000-0000-0000-0000-000000000001',
      id: '00000000-0000-0000-0000-000000000002',
      tenantId: '00000000-0000-0000-0000-000000000003',
      requiredResourceAccess: scopes
    });
    sinon.stub(entraApp, 'grantAdminConsent').resolves();
    auth.connection.accessTokens[auth.defaultResource] = {
      accessToken: 'abc',
      expiresOn: new Date().toString()
    };

    await command.action(logger, { options: { verbose: true } });

    const expected: SettingNames = {
      clientId: '00000000-0000-0000-0000-000000000001',
      tenantId: '00000000-0000-0000-0000-000000000003'
    };
    const deleted: SettingNames = {
      clientSecret: '',
      clientCertificateFile: '',
      clientCertificateBase64Encoded: '',
      clientCertificatePassword: ''
    };
    assert.deepEqual(createAppRegistrationSpy.getCall(0).args[0].apis, scopes);
    Object.keys(expected).forEach(setting => {
      assert(configSetSpy.calledWith(setting, (expected as any)[setting]), `Incorrect setting for ${setting}`);
    });
    Object.keys(deleted).forEach(setting => {
      assert(configDeleteSpy.calledWith(setting), `Not deleted ${setting}`);
    });
  });

  it(`doesn't create a new Entra app when creation not confirmed`, async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Scripting;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Proficient;
        case 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?':
          return EntraAppConfig.Create;
        case 'What scopes should the new app registration have?':
          return NewEntraAppScopes.Minimal;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return true;
        case 'CLI for Microsoft 365 will now sign in to your Microsoft 365 tenant as Microsoft Azure CLI to create a new app registration. Continue?':
          return false;
        default: //summary
          return true;
      }
    });
    const clearConnectionInfoSpy = sinon.stub(auth, 'clearConnectionInfo').resolves();

    await assert.rejects(async () => await command.action(logger, { options: {} }));
    assert(clearConnectionInfoSpy.notCalled);
  });

  it('outputs settings to configure to console in debug mode', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Interactively;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Beginner;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return false;
        default: //summary
          return true;
      }
    });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = HelpMode.Full;

    await command.action(logger, { options: { debug: true } });

    assert(loggerLogToStderrSpy.calledWith(JSON.stringify(expected, null, 2)));
  });

  it('logs configured settings when used interactively', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return CliUsageMode.Interactively;
        case 'How experienced are you in using the CLI?':
          return CliExperience.Beginner;
        default:
          return '';
      }
    });
    sinon.stub(cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig): Promise<boolean> => {
      switch (config.message) {
        case 'Are you going to use the CLI in PowerShell?':
          return false;
        default: //summary
          return true;
      }
    });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = HelpMode.Full;

    await command.action(logger, { options: {} });

    for (const [key, value] of Object.entries(expected)) {
      assert(loggerLogToStderrSpy.calledWith(formatting.getStatus(CheckStatus.Success, `${key}: ${value}`)), `Expected ${key} to be set to ${value}`);
    }
  });

  it('in the confirmation message lists all settings and their values', async () => {
    const preferences: Preferences = {
      experience: CliExperience.Beginner,
      usageMode: CliUsageMode.Interactively,
      usedInPowerShell: false
    };
    const settings = (command as any).getSettings(preferences);
    const actual = (command as any).getSummaryMessage(preferences);

    for (const [key, value] of Object.entries(settings)) {
      assert(actual.indexOf(`- ${key}: ${value}`) > -1, `Expected ${key} to be set to ${value}`);
    }
  });

  it('fails validation when both interactive and scripting options specified', async () => {
    const actual = await command.validate({
      options: {
        interactive: true,
        scripting: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when no options specified', async () => {
    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when interactive option specified', async () => {
    const actual = await command.validate({
      options: {
        interactive: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when scripting option specified', async () => {
    const actual = await command.validate({
      options: {
        scripting: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
