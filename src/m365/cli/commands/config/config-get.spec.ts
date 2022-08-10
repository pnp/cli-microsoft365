import * as assert from 'assert';
import * as sinon from 'sinon';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import { settingsNames } from '../../../../settingsNames';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./config-get');

describe(commands.CONFIG_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
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
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore(Cli.getInstance().config.get);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`gets value of the specified property`, (done) => {
    const config = Cli.getInstance().config;
    sinon.stub(config, 'get').callsFake(_ => 'json');
    command.action(logger, { options: { key: settingsNames.output } }, () => {
      try {
        assert(loggerSpy.calledWith('json'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`returns undefined if the specified setting is not configured`, (done) => {
    const config = Cli.getInstance().config;
    sinon.stub(config, 'get').callsFake(_ => undefined);
    command.action(logger, { options: { key: settingsNames.output } }, () => {
      try {
        assert(loggerSpy.calledWith(undefined));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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