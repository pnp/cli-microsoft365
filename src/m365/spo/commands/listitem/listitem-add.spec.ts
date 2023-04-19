import * as os from 'os';
import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import commands from '../../commands';
const command: Command = require('./listitem-add');

describe(commands.LISTITEM_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let ensureFolderStub: sinon.SinonStub;
  const listUrl = 'sites/project-x/documents';
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const expectedTitle = `List Item 1`;

  const expectedId = 147;
  let actualId = 0;

  const expectedContentType = 'Item';
  let actualContentType = '';

  const postFakes = async (opts: any) => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('AddValidateUpdateItemUsingPath') > -1) {
        const bodyString = JSON.stringify(opts.data);
        const ctMatch = bodyString.match(/\"?FieldName\"?:\s*\"?ContentType\"?,\s*\"?FieldValue\"?:\s*\"?(\w*)\"?/i);
        actualContentType = ctMatch ? ctMatch[1] : "";
        if (bodyString.indexOf("fail adding me") > -1) { return Promise.resolve({ value: [{ ErrorMessage: 'failed updating', 'FieldName': 'Title', 'HasException': true }] }); }
        return { value: [{ FieldName: "Id", FieldValue: expectedId, HasException: false }] };
      }
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/AddValidateUpdateItemUsingPath()`) {
      const bodyString = JSON.stringify(opts.data);
      const ctMatch = bodyString.match(/\"?FieldName\"?:\s*\"?ContentType\"?,\s*\"?FieldValue\"?:\s*\"?(\w*)\"?/i);
      actualContentType = ctMatch ? ctMatch[1] : "";
      if (bodyString.indexOf("fail adding me") > -1) { return Promise.resolve({ value: [] }); }
      return { value: [{ FieldName: "Id", FieldValue: expectedId }] };
    }
    throw 'Invalid request';
  };

  const getFakes = async (opts: any) => {
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
        "ID": actualId,
        "Modified": "2018-03-15T10:43:10Z",
        "Title": expectedTitle
      };
    }
    throw 'Invalid request';
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    ensureFolderStub = sinon.stub(spo, 'ensureFolder').resolves();
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.ensureFolder,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_ADD), true);
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

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('fails validation if listTitle and listId option not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listTitle: 'Demo List' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails to create a list item when \'fail me\' values are used', async () => {
    actualId = 0;

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      Title: "fail adding me"
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError(`Creating the item failed with the following errors: ${os.EOL}- Title - failed updating`));
    assert.strictEqual(actualId, 0);
  });

  it('returns listItemInstance object when list item is added with correct values', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    command.allowUnknownOptions();

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: webUrl,
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('creates list item in the list specified using ID', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listId: 'cf8c72a1-0207-40ee-aebd-fca67d20bc8a',
      webUrl: webUrl,
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('creates list item in the list specified using URL', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      verbose: true,
      listUrl: listUrl,
      webUrl: webUrl,
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });


  it('attempts to create the listitem with the contenttype of \'Item\' when content type option 0x01 is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: webUrl,
      contentType: expectedContentType,
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert(expectedContentType === actualContentType);
  });

  it('fails to create the listitem when the specified contentType doesn\'t exist in the target list', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      contentType: "Unexpected content type",
      Title: expectedTitle
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError("Specified content type 'Unexpected content type' doesn't exist on the target list"));
  });

  it('should call ensure folder when folder arg specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: {
        listTitle: 'Demo List',
        webUrl: webUrl,
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2"
      }
    });
    assert.strictEqual(ensureFolderStub.lastCall.args[0], 'https://contoso.sharepoint.com/sites/project-x');
    assert.strictEqual(ensureFolderStub.lastCall.args[1], '/sites/project-xxx/Lists/Demo%20List/InsideFolder2');
  });

  it('should call ensure folder when folder arg specified (debug)', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: {
        debug: true,
        listTitle: 'Demo List',
        webUrl: webUrl,
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3"
      }
    });
    assert.strictEqual(ensureFolderStub.lastCall.args[0], 'https://contoso.sharepoint.com/sites/project-x');
    assert.strictEqual(ensureFolderStub.lastCall.args[1], '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3');
  });

  it('should not have end \'/\' in the folder path when FolderPath.DecodedUrl ', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    const postStubs = sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: {
        debug: true,
        listTitle: 'Demo List',
        webUrl: webUrl,
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3/"
      }
    });
    const addValidateUpdateItemUsingPathRequest = postStubs.getCall(postStubs.callCount - 1).args[0];
    const info = addValidateUpdateItemUsingPathRequest.data.listItemCreateInfo;
    assert.strictEqual(info.FolderPath.DecodedUrl, '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3');
  });

  it('ignores global options when creating request data', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    const postStubs = sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        output: "text",
        listTitle: 'Demo List',
        webUrl: webUrl,
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3/"
      }
    });
    assert.deepEqual(postStubs.firstCall.args[0].data, {
      formValues: [{ FieldName: 'Title', FieldValue: 'List Item 1' }, { FieldName: 'ContentType', FieldValue: 'Item' }],
      listItemCreateInfo: { FolderPath: { DecodedUrl: '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3' } }
    });
  });
});
