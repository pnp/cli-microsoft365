import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
import command from './list-view-remove.js';

describe(commands.LIST_VIEW_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/ninja';
  const listId = '0cd891ef-afce-4e55-b836-fce03286cccf';
  const listTitle = 'Documents';
  const listUrl = '/sites/ninja/Shared Documents';
  const viewId = 'cc27a922-8224-4296-90a5-ebbc54da2e81';
  const viewTitle = 'MyView';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = cli.getCommandInfo(command);
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
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_VIEW_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: viewId, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: '12345', id: viewId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listTitle: listTitle, title: viewTitle } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified view from list by id and listTitle when force option not passed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        id: viewId
      }
    });

    assert(confirmationStub.calledOnce);
  });

  it('prompts before removing the specified view from list by title and listId when force option not passed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listId: listId,
        title: listTitle
      }
    });

    assert(confirmationStub.calledOnce);
  });

  it('prompts before removing the specified view from list by title and listUrl when force option not passed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listUrl: listUrl,
        title: listTitle
      }
    });

    assert(confirmationStub.calledOnce);
  });

  it('aborts removing view from list when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        id: viewId
      }
    });

    assert(postSpy.notCalled);
  });

  it('removes view from the list using id and listUrl when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')/views(guid'${formatting.encodeQueryParameter(viewId)}')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        listUrl: listUrl,
        id: viewId
      }
    });
  });

  it('removes view from the list using id and listId when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/views(guid'${formatting.encodeQueryParameter(viewId)}')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        listId: listId,
        id: viewId,
        force: true
      }
    });
  });

  it('removes view from the list using id and listTitle when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views(guid'${formatting.encodeQueryParameter(viewId)}')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        id: viewId,
        force: true
      }
    });
  });

  it('removes view from the list using title and listUrl when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listUrl: listUrl,
        title: viewTitle,
        force: true
      }
    });
  });

  it('removes view from the list using title and listId when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        listId: listId,
        title: viewTitle,
        force: true
      }
    });
  });

  it('removes view from the list using title and listTitle when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        title: viewTitle,
        force: true
      }
    });
  });

  it('correctly handles error when removing view from the list', async () => {
    const errorMessage = 'request rejected';
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: errorMessage
          }
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        listTitle: listTitle,
        title: viewTitle,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
