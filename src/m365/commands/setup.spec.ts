import * as assert from 'assert';
import * as sinon from 'sinon';
import { Cli } from '../../cli/Cli';
import { CommandInfo } from '../../cli/CommandInfo';
import { Logger } from '../../cli/Logger';
import Command from '../../Command';
import { telemetry } from '../../telemetry';
import { CheckStatus, formatting } from '../../utils/formatting';
import { pid } from '../../utils/pid';
import { session } from '../../utils/session';
import { sinonUtil } from '../../utils/sinonUtil';
import commands from './commands';
import { SettingNames } from './setup';
import { interactivePreset, powerShellPreset, scriptingPreset } from './setupPresets';
const command: Command = require('./setup');

describe(commands.SETUP, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).answers = {};
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).configureSettings,
      Cli.getInstance().config.set,
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
    (command as any).answers = {
      usageMode: 'Interactively',
      experience: 'Beginner',
      summary: true
    };
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for interactive, proficient', async () => {
    (command as any).answers = {
      usageMode: 'Interactively',
      experience: 'Proficient',
      summary: true
    };
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'options';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, non-PowerShell, beginner', async () => {
    (command as any).answers = {
      usageMode: 'Scripting',
      usedInPowerShell: false,
      experience: 'Beginner',
      summary: true
    };
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, PowerShell, beginner', async () => {
    (command as any).answers = {
      usageMode: 'Scripting',
      usedInPowerShell: true,
      experience: 'Beginner',
      summary: true
    };
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
    (command as any).answers = {
      usageMode: 'Scripting',
      usedInPowerShell: false,
      experience: 'Proficient',
      summary: true
    };
    const configureSettingsStub = sinon.stub(command as any, 'configureSettings').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    expected.helpMode = 'options';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, PowerShell, proficient', async () => {
    (command as any).answers = {
      usageMode: 'Scripting',
      usedInPowerShell: true,
      experience: 'Proficient',
      summary: true
    };
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
    (command as any).answers = {
      usageMode: 'Scripting',
      usedInPowerShell: false,
      experience: 'Beginner',
      summary: false
    };
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
    (command as any).answers = {
      usageMode: 'Interactively',
      experience: 'Beginner',
      summary: true
    };
    sinon.stub(Cli.getInstance().config, 'set').callsFake(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: { debug: true } });

    assert(loggerLogToStderrSpy.calledWith(JSON.stringify(expected, null, 2)));
  });

  it('logs configured settings when used interactively', async () => {
    (command as any).answers = {
      usageMode: 'Interactively',
      experience: 'Beginner',
      summary: true
    };
    sinon.stub(Cli.getInstance().config, 'set').callsFake(() => { });

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
