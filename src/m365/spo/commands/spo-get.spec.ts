import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command, { options } from './spo-get.js';

describe(commands.GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    auth.connection.spoUrl = undefined;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets SPO URL when no URL was get previously', async () => {
    auth.connection.spoUrl = undefined;

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'json',
        debug: true
      })
    });
    assert(loggerLogSpy.calledWith({
      SpoUrl: ''
    }));
  });

  it('gets SPO URL when other URL was get previously', async () => {
    auth.connection.spoUrl = 'https://northwind.sharepoint.com';

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'json',
        debug: true
      })
    });
    assert(loggerLogSpy.calledWith({
      SpoUrl: 'https://northwind.sharepoint.com'
    }));
  });

  it('throws error when trying to get SPO URL when not logged in to M365', async () => {
    auth.connection.active = false;

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('Log in to Microsoft 365 first'));
    assert.strictEqual(auth.connection.spoUrl, undefined);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });
});
