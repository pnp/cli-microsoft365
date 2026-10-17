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
import command, { options } from './list-contenttype-remove.js';

describe(commands.LIST_CONTENTTYPE_REMOVE, () => {
  const webUrl: string = 'https://contoso.sharepoint.com';
  const listId: string = 'dfddade1-4729-428d-881e-7fedf3cae50d';
  const listTitle: string = 'Documents';
  const contentTypeId: string = '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A';

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
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
    assert.strictEqual(command.name, commands.LIST_CONTENTTYPE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing content type from list when confirmation argument not passed (listId)', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listId: listId,
        webUrl: webUrl,
        id: contentTypeId
      })
    });

    assert(promptIssued);
  });

  it('prompts before removing content type from list when confirmation argument not passed (listTitle)', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: listTitle,
        webUrl: webUrl,
        id: contentTypeId
      })
    });

    assert(promptIssued);
  });

  it('aborts removing content type from list when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listId: listId,
        webUrl: webUrl,
        id: contentTypeId
      })
    });
  });

  it('removes content type from list when listId option is passed and prompt confirmed (debug)', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists(guid\'dfddade1-4729-428d-881e-7fedf3cae50d\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
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
        listId: listId,
        webUrl: webUrl,
        id: contentTypeId
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('removes content type from list when listTitle option is passed and prompt confirmed (debug)', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists/getByTitle(\'Documents\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
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
        listTitle: listTitle,
        webUrl: webUrl,
        id: contentTypeId
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('removes content type from list when listUrl option is passed and prompt confirmed (debug)', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetList(\'%2Fsites%2Fdocuments\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
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
        listUrl: 'sites/documents',
        webUrl: webUrl,
        id: contentTypeId
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('removes content type from list when listUrl option is passed and prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetList(\'%2Fsites%2Fdocuments\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
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
        listUrl: 'sites/documents',
        webUrl: webUrl,
        id: contentTypeId
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('removes content type from list when listId option is passed and prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists(guid\'dfddade1-4729-428d-881e-7fedf3cae50d\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
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
        listId: listId,
        webUrl: webUrl,
        id: contentTypeId
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('removes content type from list when listTitle option is passed and prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists/getByTitle(\'Documents\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
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
        listTitle: listTitle,
        webUrl: webUrl,
        id: contentTypeId
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('command correctly handles list get reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists/getByTitle(\'Documents\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: listTitle,
        webUrl: webUrl,
        id: contentTypeId,
        force: true
      })
    }), new CommandError(err));
  });

  it('uses correct API url when listTitle option is passed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists/getByTitle(\'Documents\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: listTitle,
        webUrl: webUrl,
        id: contentTypeId,
        force: true
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('uses correct API url when listId option is passed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists(guid\'dfddade1-4729-428d-881e-7fedf3cae50d\')/ContentTypes(\'0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A\')') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listId: listId,
        webUrl: webUrl,
        id: contentTypeId,
        force: true
      })
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });

  it('fails validation if neither listId nor listTitle nor listUrl are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', contentTypeId: '0x0120' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if all of the list properties are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', listUrl: 'sites/documents', contentTypeId: '0x0120' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both listId and listTitle options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '0x0120' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '0x0120' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '0x0120' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '12345', id: '0x0120' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `12345 is not a valid GUID`);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '0x0120' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if the listTitle option is passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', id: '0x0120' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', id: '0x0120' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the contentTypeId option is not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', id: '0x0120',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
