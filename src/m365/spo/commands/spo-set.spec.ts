import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command, { options } from './spo-set.js';

describe(commands.SET, () => {
  let log: string[];
  let logger: Logger;
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
  });

  afterEach(() => {
    auth.connection.spoUrl = undefined;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets SPO URL when no URL was set previously', async () => {
    auth.connection.spoUrl = undefined;

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com' }) });
    assert.strictEqual(auth.connection.spoUrl, 'https://contoso.sharepoint.com');
  });

  it('sets SPO URL when other URL was set previously', async () => {
    auth.connection.spoUrl = 'https://northwind.sharepoint.com';

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com' }) });
    assert.strictEqual(auth.connection.spoUrl, 'https://contoso.sharepoint.com');
  });

  it('trims trailing slashes from the URL', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/' }) });
    assert.strictEqual(auth.connection.spoUrl, 'https://contoso.sharepoint.com');
  });

  it('throws error when trying to set SPO URL when not logged in to M365', async () => {
    auth.connection.active = false;

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com' }) }), new CommandError('Log in to Microsoft 365 first'));
    assert.strictEqual(auth.connection.spoUrl, undefined);
  });

  it('throws error when setting the password fails', async () => {
    auth.connection.active = true;
    sinonUtil.restore(auth.storeConnectionInfo);
    sinon.stub(auth, 'storeConnectionInfo').rejects(new Error('An error has occurred while setting the password'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com' }) }), new CommandError('An error has occurred while setting the password'));
  });

  it('fails validation if url is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when the url is a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/team-a' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      url: 'https://contoso.sharepoint.com',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
