import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { CommandError } from '../../../Command.js';
import { Cli } from '../../../cli/Cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command from './spo-get.js';

describe(commands.GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
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
  });

  afterEach(() => {
    auth.service.spoUrl = undefined;
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets SPO URL when no URL was get previously', async () => {
    auth.service.spoUrl = undefined;

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      SpoUrl: ''
    }));
  });

  it('gets SPO URL when other URL was get previously', async () => {
    auth.service.spoUrl = 'https://northwind.sharepoint.com';

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      SpoUrl: 'https://northwind.sharepoint.com'
    }));
  });

  it('throws error when trying to get SPO URL when not logged in to M365', async () => {
    auth.service.active = false;

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Log in to Microsoft 365 first'));
    assert.strictEqual(auth.service.spoUrl, undefined);
  });

  it('Contains the correct options', () => {
    const options = command.options;
    let containsOutputOption = false;
    let containsVerboseOption = false;
    let containsDebugOption = false;
    let containsQueryOption = false;

    options.forEach(o => {
      if (o.option.indexOf('--output') > -1) {
        containsOutputOption = true;
      }
      else if (o.option.indexOf('--verbose') > -1) {
        containsVerboseOption = true;
      }
      else if (o.option.indexOf('--debug') > -1) {
        containsDebugOption = true;
      }
      else if (o.option.indexOf('--query') > -1) {
        containsQueryOption = true;
      }
    });

    assert(options.length === 4, "Wrong amount of options returned");
    assert(containsOutputOption, "Output option not available");
    assert(containsVerboseOption, "Verbose option not available");
    assert(containsDebugOption, "Debug option not available");
    assert(containsQueryOption, "Query option not available");
  });

  it('passes validation without any extra options', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
