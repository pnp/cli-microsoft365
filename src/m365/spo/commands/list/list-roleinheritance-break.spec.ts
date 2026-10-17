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
import command, { options } from './list-roleinheritance-break.js';

describe(commands.LIST_ROLEINHERITANCE_BREAK, () => {
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
    commandOptionsSchema = commandInfo.command.getSchemaToParse()! as typeof options;
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
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_ROLEINHERITANCE_BREAK);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'foo is not a valid SharePoint Online site URL.');
  });

  it('fails validation if listId and listTitle are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listUrl are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listUrl: '/sites/Documents' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listTitle and listUrl are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', listUrl: '/sites/Documents' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither listTitle nor listId or listUrl is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '12345' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `12345 is not a valid GUID`);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, true);
  });

  it('break role inheritance on list by title', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists/getByTitle(\'test\')/breakroleinheritance(true)') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        force: true
      })
    });
  });

  it('break role inheritance on list by title and clear all permissions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists/getByTitle(\'test\')/breakroleinheritance(false)') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        clearExistingPermissions: true,
        force: true
      })
    });
  });

  it('reset role inheritance on list by list url and break all the permissions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetList(\'%2Fsites%2Fdocuments\')/breakroleinheritance(false)') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listUrl: 'sites/documents',
        clearExistingPermissions: true,
        force: true
      })
    });
  });

  it('break role inheritance on list by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists(guid\'202b8199-b9de-43fd-9737-7f213f51c991\')/breakroleinheritance(true)') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '202b8199-b9de-43fd-9737-7f213f51c991',
        force: true
      })
    });
  });

  it('break role inheritance on list by id and clear all permissions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists(guid\'202b8199-b9de-43fd-9737-7f213f51c991\')/breakroleinheritance(false)') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '202b8199-b9de-43fd-9737-7f213f51c991',
        clearExistingPermissions: true,
        force: true
      })
    });
  });

  it('list role inheritance break command handles reject request correctly', async () => {
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
      if ((opts.url as string).indexOf('/_api/web/lists/getByTitle(\'test\')/breakroleinheritance(true)') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        force: true
      })
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('aborts breaking role inheritance when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    const postSpy = sinon.spy(request, 'post');
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test'
      })
    });
    assert(postSpy.notCalled);
  });

  it('prompts before breaking role inheritance when confirmation argument not passed (Title)', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test'
      })
    });
    assert(promptIssued);
  });

  it('prompts before breaking role inheritance when confirmation argument not passed (id)', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '202b8199-b9de-43fd-9737-7f213f51c991'
      })
    });
    assert(promptIssued);
  });

  it('break role inheritance when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getByTitle(\'test\')/breakroleinheritance(true)') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test'
      })
    });
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
