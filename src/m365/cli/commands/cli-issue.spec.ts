import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../telemetry';
import { Cli } from '../../../cli/Cli';
import { CommandInfo } from '../../../cli/CommandInfo';
import { Logger } from '../../../cli/Logger';
import Command from '../../../Command';
import { pid } from '../../../utils/pid';
import { sinonUtil } from '../../../utils/sinonUtil';
import commands from '../commands';
import Sinon = require('sinon');

const command: Command = require('./cli-issue');

describe(commands.ISSUE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let openBrowserSpy: Sinon.SinonSpy;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    (command as any).open = () => { };
    openBrowserSpy = sinon.spy(command as any, 'openBrowser');
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
  });

  afterEach(() => {
    openBrowserSpy.resetHistory();
  });

  after(() => {
    sinonUtil.restore([
      telemetry.trackEvent,
      pid.getProcessName
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ISSUE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('accepts Bug issue Type', async () => {
    const actual = await command.validate({ options: { type: 'bug' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts Command issue Type', async () => {
    const actual = await command.validate({ options: { type: 'command' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts Sample issue Type', async () => {
    const actual = await command.validate({ options: { type: 'sample' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid issue type', async () => {
    const type = 'foo';
    const actual = await command.validate({ options: { type: type } }, commandInfo);
    assert.strictEqual(actual, `${type} is not a valid Issue type. Allowed values are bug, command, sample`);
  });

  it('Opens URL for a command (debug)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        type: 'command'
      }
    } as any);
    openBrowserSpy.calledWith("https://aka.ms/cli-m365/new-command");
  });

  it('Opens URL for a bug (debug)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        type: 'bug'
      }
    } as any);
    openBrowserSpy.calledWith("https://aka.ms/cli-m365/bug");
  });

  it('Opens URL for a sample (debug)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        type: 'sample'
      }
    } as any);
    openBrowserSpy.calledWith("https://aka.ms/cli-m365/new-sample-script");
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
