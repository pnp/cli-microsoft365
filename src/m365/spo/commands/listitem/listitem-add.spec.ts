import assert from 'assert';
import os from 'os';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command, { options } from './listitem-add.js';

describe(commands.LISTITEM_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let ensureFolderStub: sinon.SinonStub;
  const listUrl = 'sites/project-x/documents';
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const expectedTitle = `List Item 1`;

  const expectedId = 147;
  let actualId = 0;

  const expectedContentType = 'Item';
  let actualContentType = '';

  const postFakes = async (opts: any): Promise<any> => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('AddValidateUpdateItemUsingPath') > -1) {
        const bodyString = JSON.stringify(opts.data);
        const ctMatch = bodyString.match(/"?FieldName"?:\s*"?ContentType"?,\s*"?FieldValue"?:\s*"?(\w*)"?/i);
        actualContentType = ctMatch ? ctMatch[1] : "";
        if (bodyString.indexOf("fail adding me") > -1) { return Promise.resolve({ value: [{ ErrorMessage: 'failed updating', 'FieldName': 'Title', 'HasException': true }] }); }
        return { value: [{ FieldName: "Id", FieldValue: expectedId, HasException: false }] };
      }
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/AddValidateUpdateItemUsingPath()`) {
      const bodyString = JSON.stringify(opts.data);
      const ctMatch = bodyString.match(/"?FieldName"?:\s*"?ContentType"?,\s*"?FieldValue"?:\s*"?(\w*)"?/i);
      actualContentType = ctMatch ? ctMatch[1] : "";
      if (bodyString.indexOf("fail adding me") > -1) { return Promise.resolve({ value: [] }); }
      return { value: [{ FieldName: "Id", FieldValue: expectedId }] };
    }
    throw 'Invalid request';
  };

  const getFakes = async (opts: any): Promise<any> => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('contenttypes') > -1) {
        return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
      }
      if ((opts.url as string).indexOf('rootFolder') > -1) {
        return { ServerRelativeUrl: '/sites/project-xxx/Lists/Demo%20List' };
      }
      if ((opts.url as string).indexOf('/items(') > -1) {
        actualId = parseInt(opts.url.match(/\/items\((\d+)\)/i)[1]);
        return {
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-03-15T10:43:10Z",
          "EditorId": 3,
          "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
          "Id": actualId,
          "ID": actualId,
          "Modified": "2018-03-15T10:43:10Z",
          "Title": expectedTitle
        };
      }
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/contenttypes?$select=Name,Id`) {
      return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(147)`) {
      actualId = parseInt(opts.url.match(/\/items\((\d+)\)/i)[1]);
      return {
        "Attachments": false,
        "AuthorId": 3,
        "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
        "Created": "2018-03-15T10:43:10Z",
        "EditorId": 3,
        "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
        "Id": actualId,
        "ID": actualId,
        "Modified": "2018-03-15T10:43:10Z",
        "Title": expectedTitle
      };
    }
    throw 'Invalid request';
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    ensureFolderStub = sinon.stub(spo, 'ensureFolder').resolves();
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
    sinonUtil.restore([
      request.post,
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('allows unknown options', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  it('fails validation if listTitle and listId option not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Use one of the following options: listId, listTitle, listUrl.');
  });

  it('fails validation if listTitle and listId are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Use one of the following options: listId, listTitle, listUrl.');
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listTitle: 'Demo List' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: 'foo' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'foo in option listId is not a valid GUID');
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, true);
  });

  it('fails to create a list item when \'fail me\' values are used', async () => {
    actualId = 0;

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      Title: "fail adding me"
    };

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse(options) }), new CommandError(`Creating the item failed with the following errors: ${os.EOL}- Title - failed updating`));
    assert.strictEqual(actualId, 0);
  });

  it('returns listItemInstance object when list item is added with correct values', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: webUrl,
      Title: expectedTitle
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert.strictEqual(actualId, expectedId);
  });

  it('creates list item in the list specified using ID', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options = {
      listId: 'cf8c72a1-0207-40ee-aebd-fca67d20bc8a',
      webUrl: webUrl,
      Title: expectedTitle
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert.strictEqual(actualId, expectedId);
  });

  it('creates list item in the list specified using URL', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options = {
      verbose: true,
      listUrl: listUrl,
      webUrl: webUrl,
      Title: expectedTitle
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert.strictEqual(actualId, expectedId);
  });


  it('attempts to create the listitem with the contenttype of \'Item\' when content type option 0x01 is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: webUrl,
      contentType: expectedContentType,
      Title: expectedTitle
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(expectedContentType === actualContentType);
  });

  it('fails to create the listitem when the specified contentType doesn\'t exist in the target list', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      contentType: "Unexpected content type",
      Title: expectedTitle
    };

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse(options) }), new CommandError("Specified content type 'Unexpected content type' doesn't exist on the target list"));
  });

  it('should call ensure folder when folder arg specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: 'Demo List',
        webUrl: webUrl,
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2"
      })
    });
    assert.strictEqual(ensureFolderStub.lastCall.args[0], 'https://contoso.sharepoint.com/sites/project-x');
    assert.strictEqual(ensureFolderStub.lastCall.args[1], '/sites/project-xxx/Lists/Demo%20List/InsideFolder2');
  });

  it('should call ensure folder when folder arg specified (debug)', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: 'Demo List',
        webUrl: webUrl,
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3"
      })
    });
    assert.strictEqual(ensureFolderStub.lastCall.args[0], 'https://contoso.sharepoint.com/sites/project-x');
    assert.strictEqual(ensureFolderStub.lastCall.args[1], '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3');
  });

  it('should not have end \'/\' in the folder path when FolderPath.DecodedUrl ', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    const postStubs = sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: 'Demo List',
        webUrl: webUrl,
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3/"
      })
    });
    const addValidateUpdateItemUsingPathRequest = postStubs.getCall(postStubs.callCount - 1).args[0];
    const info = addValidateUpdateItemUsingPathRequest.data.listItemCreateInfo;
    assert.strictEqual(info.FolderPath.DecodedUrl, '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3');
  });

  it('ignores global options when creating request data', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    const postStubs = sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        verbose: true,
        output: "text",
        listTitle: 'Demo List',
        webUrl: webUrl,
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3/"
      })
    });
    assert.deepEqual(postStubs.firstCall.args[0].data, {
      formValues: [{ FieldName: 'Title', FieldValue: 'List Item 1' }, { FieldName: 'ContentType', FieldValue: 'Item' }],
      listItemCreateInfo: { FolderPath: { DecodedUrl: '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3' } }
    });
  });
});
