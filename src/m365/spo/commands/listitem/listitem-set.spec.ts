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
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './listitem-set.js';
import { settingsNames } from '../../../../settingsNames.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';

describe(commands.LISTITEM_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;

  const webUrl = 'https://contoso.sharepoint.com/sites/project-w';
  const listUrl = '/sites/project-x/lists/TestList';
  const listId = '9befab64-10fa-4a1a-88ad-200629d5306a';
  const listItemResponse = {
    "FileSystemObjectType": 0,
    "Id": 1,
    "ServerRedirectedEmbedUri": null,
    "ServerRedirectedEmbedUrl": "",
    "ContentTypeId": "0x0100A06E900513958643B1CBA90ACB57A4C70088931AAA291F244FA07D46D3B40AD0F1",
    "Title": "Test",
    "OData__ColorTag": null,
    "ComplianceAssetId": null,
    "ID": 1,
    "Modified": new Date("2023-10-30T15:36:11Z"),
    "Created": new Date("2023-10-16T12:40:57Z"),
    "AuthorId": 11,
    "EditorId": 11,
    "OData__UIVersionString": "3.0",
    "Attachments": false,
    "GUID": "fe213cee-4c05-4de8-a306-f8a5f0923d5a",
    "RoleAssignments": []
  };

  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const expectedTitle = `List Item 1`;
  const expectedContentType = 'Item';

  const getFakes = async (opts: any) => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('contenttypes') > -1) {
        return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
      }
    }

    if (opts.url === `https://contoso.sharepoint.com/sites/project-w/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=Id`) {
      return { Id: "f64041f2-9818-4b67-92ff-3bc5dbbef27e" };
    }

    if (opts.url === `https://contoso.sharepoint.com/sites/project-w/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/contenttypes?$select=Name,Id`) {
      return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
    }

    if (opts.url === `https://contoso.sharepoint.com/sites/project-y/_api/web/lists(guid'${listId}')/contenttypes?$select=Name,Id`) {
      return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
    }

    throw 'Invalid request';
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      spo.updateListItem,
      spo.systemUpdateListItem,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_SET);
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle, listId and listUrl are specified together', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(spo, 'updateListItem').rejects(new Error(`Updating the items has failed with the following errors: ${os.EOL}- Title - failed updating`));
    const options: any = {
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      Title: "fail updating me"
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError(`Updating the items has failed with the following errors: ${os.EOL}- Title - failed updating`));
  });

  it('returns listItemInstance object when list item is updated with correct values', async () => {
    sinon.stub(spo, 'updateListItem').resolves(listItemResponse);
    command.allowUnknownOptions();

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(listItemResponse));
  });

  it('returns listItemInstance object when list item in list retrieved by URL is updated with correct values', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(spo, 'systemUpdateListItem').resolves(listItemResponse);
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
    assert(loggerLogSpy.calledWith(listItemResponse));
  });

  it('attempts to update the listitem with the contenttype of \'Item\' when content type option \'Item\' is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(spo, 'updateListItem').resolves(listItemResponse);

    const options: any = {
      listTitle: 'Demo List',
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      contentType: 'Item',
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(listItemResponse));
  });

  it('attempts to update the listitem with the contenttype of \'Item\' when content type option 0x01 is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(spo, 'updateListItem').resolves(listItemResponse);

    const options: any = {
      debug: true,
      listId: listId,
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      contentType: expectedContentType,
      Title: expectedTitle
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(listItemResponse));
  });

  it('fails to update the listitem when the specified contentType doesn\'t exist in the target list', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

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
    sinon.stub(spo, 'systemUpdateListItem').resolves(listItemResponse);

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      id: 147,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      Title: expectedTitle,
      systemUpdate: true
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(listItemResponse));
  });

  it('fails to update the list item when systemUpdate parameter is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(spo, 'systemUpdateListItem').rejects(new Error('Error occurred in systemUpdate operation - ErrorMessage": "systemUpdate error"}'));

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
  });

  it('should ignore global options when creating request data', async () => {
    sinon.stub(spo, 'updateListItem').resolves(listItemResponse);

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
    assert(loggerLogSpy.calledWith(listItemResponse));
  });
});
