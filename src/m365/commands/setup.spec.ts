import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../../cli/cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { telemetry } from '../../telemetry.js';
import { CheckStatus, formatting } from '../../utils/formatting.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import commands from './commands.js';
import command, { SettingNames } from './setup.js';
import { interactivePreset, powerShellPreset, scriptingPreset } from './setupPresets.js';
import { ConfirmationConfig, SelectionConfig } from '../../utils/prompt.js';

describe(commands.SETUP, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = cli.getCommandInfo(command);
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
      cli.promptForSelection,
      cli.getConfig().set,
      pid.isPowerShell
    ]);
  });

  after(() => {
    sinonUtil.restore([
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
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
          return 'Interactively';
        case 'How experienced are you in using the CLI?':
          return 'Beginner';
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

    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for interactive, proficient', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return 'Interactively';
        case 'How experienced are you in using the CLI?':
          return 'Proficient';
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
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'options';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, non-PowerShell, beginner', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return 'Scripting';
        case 'How experienced are you in using the CLI?':
          return 'Beginner';
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
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, PowerShell, beginner', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return 'Scripting';
        case 'How experienced are you in using the CLI?':
          return 'Beginner';
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
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    Object.assign(expected, powerShellPreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, non-PowerShell, proficient', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return 'Scripting';
        case 'How experienced are you in using the CLI?':
          return 'Proficient';
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
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    expected.helpMode = 'options';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, PowerShell, proficient', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return 'Scripting';
        case 'How experienced are you in using the CLI?':
          return 'Proficient';
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
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    Object.assign(expected, powerShellPreset);
    expected.helpMode = 'options';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it(`doesn't apply settings when not confirmed`, async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return 'Scripting';
        case 'How experienced are you in using the CLI?':
          return 'Beginner';
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
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    (command as any).settings = expected;

    await command.action(logger, { options: { interactive: true } });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, non-PowerShell via option', async () => {
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    (command as any).settings = expected;

    await command.action(logger, { options: { scripting: true } });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for interactive, PowerShell via option', async () => {
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });
    sinon.stub(pid, 'isPowerShell').callsFake(() => true);

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    Object.assign(expected, powerShellPreset);
    (command as any).settings = expected;

    await command.action(logger, { options: { interactive: true } });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, PowerShell via option', async () => {
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });
    sinon.stub(pid, 'isPowerShell').callsFake(() => true);

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    Object.assign(expected, powerShellPreset);
    (command as any).settings = expected;

    await command.action(logger, { options: { scripting: true } });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('outputs settings to configure to console in debug mode', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return 'Interactively';
        case 'How experienced are you in using the CLI?':
          return 'Beginner';
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
    sinon.stub(cli.getConfig(), 'set').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: { debug: true } });

    assert(loggerLogToStderrSpy.calledWith(JSON.stringify(expected, null, 2)));
  });

  it('logs configured settings when used interactively', async () => {
    sinon.stub(cli, 'promptForSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      switch (config.message) {
        case 'How do you plan to use the CLI?':
          return 'Interactively';
        case 'How experienced are you in using the CLI?':
          return 'Beginner';
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
    sinon.stub(cli.getConfig(), 'set').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    for (const [key, value] of Object.entries(expected)) {
      assert(loggerLogToStderrSpy.calledWith(formatting.getStatus(CheckStatus.Success, `${key}: ${value}`)), `Expected ${key} to be set to ${value}`);
    }
  });

  it('in the confirmation message lists all settings and their values', async () => {
    const settings: SettingNames = {};
    Object.assign(settings, interactivePreset);
    settings.helpMode = 'full';
    const actual = (command as any).getSummaryMessage(settings);

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
