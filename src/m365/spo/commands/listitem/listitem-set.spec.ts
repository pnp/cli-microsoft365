import * as assert from 'assert';
import * as sinon from 'sinon';
import * as os from 'os';
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
const command: Command = require('./listitem-set');

describe(commands.LISTITEM_SET, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com/sites/project-w';
  const listUrl = '/sites/project-x/lists/TestList';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

  const expectedTitle = `List Item 1`;

  const expectedId = 147;
  let actualId = 0;

  const expectedContentType = 'Item';
  let actualContentType = '';

  const postFakes = async (opts: any) => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('ValidateUpdateListItem') > -1) {
        const bodyString = JSON.stringify(opts.data);
        const ctMatch = bodyString.match(/\"?FieldName\"?:\s*\"?ContentType\"?,\s*\"?FieldValue\"?:\s*\"?(\w*)\"?/i);
        actualContentType = ctMatch ? ctMatch[1] : "";
        if (bodyString.indexOf("fail updating me") > -1) { return Promise.resolve({ value: [{ ErrorMessage: 'failed updating', 'FieldName': 'Title', 'HasException': true }] }); }
        return { value: [{ ItemId: expectedId }] };
      }
    }

    if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
      if (opts.data.indexOf('Name="Current"') > -1) {
        if ((opts.url as string).indexOf('rejectme.com') > -1) {
          throw 'Failed request';
        }

        if ((opts.url as string).indexOf('returnerror.com') > -1) {
          return JSON.stringify([{ "ErrorInfo": "error occurred" }]);
        }

        if (opts.url === `https://objectidentityNotFound.sharepoint.com/sites/project-y/_vti_bin/client.svc/ProcessQuery`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7618.1204",
              "ErrorInfo": null,
              "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
            },
            {
              "_ObjectType_": "SP.Web",
              "ServerRelativeUrl": "\\u002fsites\\u002fprojecty"
            }
          ]);
        }

        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7618.1204",
            "ErrorInfo": null,
            "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
          },
          {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7",
            "ServerRelativeUrl": "\\u002fsites\\u002fprojectx"
          }
        ]);
      }
      if (opts.data.indexOf('SystemUpdate') > -1) {
        if (opts.data.indexOf('systemUpdate error') > -1) {
          return 'ErrorMessage": "systemUpdate error"}';
        }
        actualId = expectedId;
        return ']SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7618.1204","ErrorInfo":null,"TraceCorrelationId":"3e3e629e-f0e9-5000-9f31-c6758b453a4a"';
      }
    }

    throw 'Invalid request';
  };

  const getFakes = async (opts: any) => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('contenttypes') > -1) {
        return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
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
          "Modified": "2018-03-15T10:52:10Z",
          "Title": expectedTitle
        };
      }
      if ((opts.url as string).indexOf(')?$select=Id') > -1) {
        return { Id: "f64041f2-9818-4b67-92ff-3bc5dbbef27e" };
      }
    }

    if (opts.url === `https://contoso.sharepoint.com/sites/project-w/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=Id`) {
      return { Id: "f64041f2-9818-4b67-92ff-3bc5dbbef27e" };
    }

    if (opts.url === `https://contoso.sharepoint.com/sites/project-w/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/contenttypes?$select=Name,Id`) {
      return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
    }

    if (opts.url === `https://contoso.sharepoint.com/sites/project-w/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(147)`) {
      actualId = parseInt(opts.url.match(/\/items\((\d+)\)/i)[1]);
      return {
        "Attachments": false,
        "AuthorId": 3,
        "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
        "Created": "2018-03-15T10:43:10Z",
        "EditorId": 3,
        "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
        "ID": actualId,
        "Modified": "2018-03-15T10:52:10Z",
        "Title": expectedTitle
      };
    }

    throw 'Invalid request';
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => {
      if (settingName === "prompt") { return false; }
      else {
        return defaultValue;
      }
    }));
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
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_SET), true);
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

  it('fails validation if listTitle, listId or listUrl option not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle, listId and listUrl are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listUrl: listUrl, id: '1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listTitle: 'Demo List', id: '1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: '1' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo', id: '1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '1' } }, commandInfo);
    assert(actual);
  });

  it('fails to update a list item when \'fail me\' values are used', async () => {
    actualId = 0;

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      Title: "fail updating me"
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError(`Updating the items has failed with the following errors: ${os.EOL}- Title - failed updating`));
    assert.strictEqual(actualId, 0);
  });

  it('returns listItemInstance object when list item is updated with correct values', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    command.allowUnknownOptions();

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item in list retrieved by URL is updated with correct values', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    command.allowUnknownOptions();

    const options: any = {
      verbose: true,
      listUrl: listUrl,
      id: 147,
      webUrl: webUrl,
      contentType: 'Item',
      Title: expectedTitle,
      systemUpdate: true
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('attempts to update the listitem with the contenttype of \'Item\' when content type option \'Item\' is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      contentType: 'Item',
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert(expectedContentType === actualContentType);
  });

  it('attempts to update the listitem with the contenttype of \'Item\' when content type option 0x01 is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      contentType: expectedContentType,
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert(expectedContentType === actualContentType);
  });

  it('fails to update the listitem when the specified contentType doesn\'t exist in the target list', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      contentType: "Unexpected content type",
      Title: expectedTitle
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError("Specified content type 'Unexpected content type' doesn't exist on the target list"));
  });

  it('successfully updates the listitem when the systemUpdate parameter is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    actualId = 0;

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      id: 147,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      Title: expectedTitle,
      systemUpdate: true
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('fails to get _ObjecttIdentity_ when the systemUpdate parameter is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    actualId = 0;

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      id: 147,
      webUrl: 'https://rejectme.com/sites/project-y',
      Title: expectedTitle,
      systemUpdate: true
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError("Failed request"));
  });

  it('fails to get _ObjecttIdentity_ when objectidentity not found', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    actualId = 0;

    const options: any = {
      debug: true,
      listTitle: 'Test List',
      id: 147,
      webUrl: 'https://objectidentityNotFound.sharepoint.com/sites/project-y',
      Title: expectedTitle,
      systemUpdate: true
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError("Cannot proceed. _ObjectIdentity_ not found"));
  });

  it('fails to get _ObjecttIdentity_ when an error is returned by the _ObjectIdentity_ CSOM request and systemUpdate parameter is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    actualId = 0;

    const options: any = {
      listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
      id: 147,
      webUrl: 'https://returnerror.com/sites/project-y',
      Title: expectedTitle,
      systemUpdate: true
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('ClientSvc unknown error'));
    assert(actualId !== expectedId);
  });

  it('fails to update the list item when systemUpdate parameter is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    actualId = 0;

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      id: 147,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      Title: "systemUpdate error",
      contentType: "Item",
      systemUpdate: true
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('Error occurred in systemUpdate operation - ErrorMessage": "systemUpdate error"}'));
    assert(actualId !== expectedId);
  });

  it('should ignore global options when creating request data', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    const postStubs = sinon.stub(request, 'post').callsFake(postFakes);

    actualId = 0;

    const options: any = {
      debug: true,
      verbose: true,
      output: "text",
      listTitle: 'Demo List',
      id: 147,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      Title: expectedTitle,
      systemUpdate: false
    };

    await command.action(logger, { options: options } as any);
    assert.deepEqual(postStubs.firstCall.args[0].data, { formValues: [{ FieldName: 'Title', FieldValue: 'List Item 1' }] });
  });
});
