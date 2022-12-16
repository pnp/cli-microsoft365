import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import { settingsNames } from '../../../../settingsNames';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./config-get');

describe(commands.CONFIG_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    commandInfo = Cli.getCommandInfo(command);
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore(Cli.getInstance().config.get);
  });

  after(() => {
    sinonUtil.restore([
      telemetry.trackEvent,
      pid.getProcessName
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`gets value of the specified property`, async () => {
    const config = Cli.getInstance().config;
    sinon.stub(config, 'get').callsFake(_ => 'json');
    await command.action(logger, { options: { key: settingsNames.output } });
    assert(loggerSpy.calledWith('json'));
  });

  it(`returns undefined if the specified setting is not configured`, async () => {
    const config = Cli.getInstance().config;
    sinon.stub(config, 'get').callsFake(_ => undefined);
    await command.action(logger, { options: { key: settingsNames.output } });
    assert(loggerSpy.calledWith(undefined));
  });

  it('supports specifying key', () => {
    const options = command.options;
    let containsOptionKey = false;
    options.forEach(o => {
      if (o.option.indexOf('--key') > -1) {
        containsOptionKey = true;
      }
    });
    assert(containsOptionKey);
  });

  it('fails validation if specified key is invalid ', async () => {
    const actual = await command.validate({ options: { key: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if setting is set to ${settingsNames.showHelpOnFailure}`, async () => {
    const actual = await command.validate({ options: { key: settingsNames.showHelpOnFailure } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
