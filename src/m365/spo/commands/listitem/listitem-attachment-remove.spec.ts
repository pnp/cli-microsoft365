import assert from 'assert';
import sinon from 'sinon';
import { telemetry } from '../../../../telemetry.js';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './listitem-attachment-remove.js';

describe(commands.LISTITEM_ATTACHMENT_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listId = '4fc5ba1e-18b7-49e0-81fe-54515cc2eede';
  const listTitle = 'Demo List';
  const listUrl = '/sites/project-x/Lists/DemoList';
  const listItemId = 147;
  const fileName = 'File1.jpg';

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean;

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
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ATTACHMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listTitle: listTitle, listItemId: listItemId, fileName: fileName });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: listTitle, listItemId: listItemId, fileName: fileName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: 'foo', listItemId: listItemId, fileName: fileName });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'foo in option listId is not a valid GUID');
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: listId, listItemId: listItemId, fileName: fileName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the specified listItemId is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: listTitle, listItemId: 'a', fileName: fileName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: listTitle, listItemId: listItemId, fileName: fileName, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('prompts before removing attachment from list item when confirmation argument not passed (listId)', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: webUrl,
        listId: listId,
        listItemId: listItemId,
        fileName: fileName
      })
    });

    assert(promptIssued);
  });

  it('prompts before removing attachment from list item when force argument not passed (listTitle)', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: webUrl,
        listTitle: listTitle,
        listItemId: listItemId,
        fileName: fileName
      })
    });

    assert(promptIssued);
  });

  it('aborts removing attachment from list item when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: webUrl,
        listId: listId,
        listItemId: listItemId,
        fileName: fileName
      })
    });
  });

  it('removes attachment from list item when listId option is passed and prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'4fc5ba1e-18b7-49e0-81fe-54515cc2eede')/items(147)/AttachmentFiles('File1.jpg')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: webUrl,
        listId: listId,
        listItemId: listItemId,
        fileName: fileName
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('removes attachment from list item when listTitle option is passed and prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Demo%20List')/items(147)/AttachmentFiles('File1.jpg')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: webUrl,
        listTitle: listTitle,
        listItemId: listItemId,
        fileName: fileName
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('removes attachment from list item when listUrl option is passed and prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('%2Fsites%2Fproject-x%2FLists%2FDemoList')/items(147)/AttachmentFiles('File1.jpg')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: webUrl,
        listUrl: listUrl,
        listItemId: listItemId,
        fileName: fileName
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('command correctly handles list get reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Demo%20List')/items(147)/AttachmentFiles('File1.jpg')`) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: webUrl,
        listTitle: listTitle,
        listItemId: listItemId,
        fileName: fileName,
        force: true
      })
    }), new CommandError(err));
  });

  it('uses correct API url when listTitle option is passed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Demo%20List')/items(147)/AttachmentFiles('File1.jpg')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: webUrl,
        listTitle: listTitle,
        listItemId: listItemId,
        fileName: fileName,
        force: true
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });
});