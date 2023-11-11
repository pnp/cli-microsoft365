import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command from './listitem-roleinheritance-break.js';

describe(commands.LISTITEM_ROLEINHERITANCE_BREAK, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
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
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ROLEINHERITANCE_BREAK);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listItemId: '4', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: '4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', listItemId: '4' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: '4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the specified list item id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listItemId: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the specified list item id is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listItemId: '4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('break role inheritance of list item with id 1 on list by list url', async () => {
    const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
    const listUrl = '/sites/project-x/lists/TestList';
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
    const listItemId = 8;

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(${listItemId})/breakroleinheritance(true)`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listUrl: listUrl,
        listItemId: listItemId
      }
    });
  });

  it('break role inheritance of list item with id 1 on list by title', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle(\'test\')/items(1)/breakroleinheritance(true)') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        listItemId: 1,
        force: true
      }
    });
  });

  it('break role inheritance of list item with id 1 on list by title and clear all permissions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle(\'test\')/items(1)/breakroleinheritance(false)') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'test',
        listItemId: 1,
        clearExistingPermissions: true,
        force: true
      }
    });
  });

  it('break role inheritance of list item with id 1 on list by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'202b8199-b9de-43fd-9737-7f213f51c991\')/items(1)/breakroleinheritance(true)') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 1,
        listId: '202b8199-b9de-43fd-9737-7f213f51c991',
        force: true
      }
    });
  });

  it('break role inheritance of list item with id 1 on list by id and clear all permissions', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'202b8199-b9de-43fd-9737-7f213f51c991\')/items(1)/breakroleinheritance(false)') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '202b8199-b9de-43fd-9737-7f213f51c991',
        listItemId: 1,
        clearExistingPermissions: true,
        force: true
      }
    });
  });

  it('list item role inheritance break command handles reject request correctly', async () => {
    const err = 'request rejected';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle(\'test\')/items(1)/breakroleinheritance(true)') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 1,
        listTitle: 'test',
        force: true
      }
    }), new CommandError(err));
  });

  it('aborts breaking role inheritance when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 8,
        listTitle: 'test'
      }
    });
    assert(postSpy.notCalled);
  });

  it('prompts before breaking role inheritance when confirmation argument not passed (Title)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 8,
        listTitle: 'test'
      }
    });


    assert(promptIssued);
  });

  it('prompts before breaking role inheritance when confirmation argument not passed (id)', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 8,
        listId: '202b8199-b9de-43fd-9737-7f213f51c991'
      }
    });


    assert(promptIssued);
  });

  it('break role inheritance of list item with id 1 on list by list url without confirmation prompt', async () => {
    const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
    const listUrl = '/sites/project-x/lists/TestList';
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
    const listItemId = 8;

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(${listItemId})/breakroleinheritance(true)`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listUrl: listUrl,
        listItemId: listItemId,
        force: true
      }
    });
  });

  it('break role inheritance when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle(\'test\')/items(8)/breakroleinheritance(true)') > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 8,
        listTitle: 'test'
      }
    });
  });
});
