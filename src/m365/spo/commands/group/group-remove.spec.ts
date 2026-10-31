import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './group-remove.js';

describe(commands.GROUP_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('deletes the group when id is passed', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true, force: true }) });
    assert(requestPostSpy.called);
  });

  it('deletes the group when name is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/mysite/_api/web/sitegroups/GetByName('Team Site Owners')?$select=Id`) {
        return {
          Id: 7
        };
      }
      throw 'Invalid request';
    });

    const requestPostSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/mysite', name: 'Team Site Owners', debug: true, force: true }) });
    assert(requestPostSpy.called);
  });

  it('aborts deleting the group when prompt is not continued', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true }) });
    assert(requestPostSpy.notCalled);
  });

  it('deletes the group when prompt is continued', async () => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true }) });
    assert(requestPostSpy.called);
  });

  it('correctly handles group remove reject request', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        throw error;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true, force: true }) }),
      new CommandError(error.error['odata.error'].message.value));
  });

  it('prompts before removing group when confirmation argument not passed (id)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ id: 7, webUrl: 'https://contoso.sharepoint.com/mysite' }) });

    assert(promptIssued);
  });

  it('prompts before removing group when confirmation argument not passed (name)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'Team Site Owners', webUrl: 'https://contoso.sharepoint.com/mysite' }) });

    assert(promptIssued);
  });

  it('fails validation if both id and name options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/mysite' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', id: 7 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/mysite', id: 7 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the id option is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/mysite', id: 'Hi' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the id option is a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/mysite', id: 7 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both id and name options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, name: 'Team Site Members' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/mysite',
      id: 7,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
