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
import command, { options } from './listitem-get.js';

describe(commands.LISTITEM_GET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listUrl = 'sites/project-x/documents';
  const listTitle = 'Demo List';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const expectedTitle = `List Item 1`;
  const expectedId = 147;
  const expectedUniqueId = 'ea093c7b-8ae6-4400-8b75-e2d01154dffc';

  let actualId = 0;

  const getFakes = async (opts: any): Promise<any> => {
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(147)/RoleAssignments?$expand=Member,RoleDefinitionBindings` ||
      opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/items(147)/RoleAssignments?$expand=Member,RoleDefinitionBindings`
    ) {
      return {
        "value": [
          {
            "Member": {
              "Id": 3,
              "IsHiddenInUI": false,
              "LoginName": "Communication site Owners",
              "Title": "Communication site Owners",
              "PrincipalType": 8,
              "AllowMembersEditMembership": false,
              "AllowRequestToJoinLeave": false,
              "AutoAcceptRequestToJoinLeave": false,
              "Description": null,
              "OnlyAllowMembersViewMembership": false,
              "OwnerTitle": "Communication site Owners",
              "RequestToJoinLeaveEmailSetting": ""
            },
            "RoleDefinitionBindings": [
              {
                "BasePermissions": {
                  "High": "2147483647",
                  "Low": "4294967295"
                },
                "Description": "Has full control.",
                "Hidden": false,
                "Id": 1073741829,
                "Name": "Full Control",
                "Order": 1,
                "RoleTypeKind": 5
              }
            ],
            "PrincipalId": 3
          }
        ]
      };
    }

    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Demo%20List')/GetItemByUniqueId(guid'ea093c7b-8ae6-4400-8b75-e2d01154dffc')?$select=`) {
      return {
        "Attachments": false,
        "AuthorId": 3,
        "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
        "Created": "2018-03-15T10:43:10Z",
        "EditorId": 3,
        "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
        "ID": 147,
        "Modified": "2018-03-15T10:43:10Z",
        "Title": expectedTitle
      };
    }

    if (opts.url.indexOf('/_api/web/lists') > -1) {
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

    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(147)?$select=${formatting.encodeQueryParameter('Title,Modified')}`) {
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
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if listTitle, listId or listUrl option not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: expectedId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listTitle, listId and listUrl are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listUrl: listUrl, id: expectedId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listTitle: 'Demo List', id: expectedId });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: expectedId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: 'foo', id: expectedId });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'foo in option listId is not a valid GUID');
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: expectedId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the uniqueId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', uniqueId: 'foo' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'foo in option uniqueId is not a valid GUID');
  });

  it('passes validation if the uniqueId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', uniqueId: expectedUniqueId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the specified id is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: 'a' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both id and uniqueId are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: expectedId, uniqueId: expectedUniqueId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: expectedId, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Unrecognized key: "unknownOption"');
  });

  it('returns listItemInstance object by id when list item is requested', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, listTitle: 'Demo List', webUrl: webUrl, id: expectedId }) });
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object by uniqueId when list item is requested', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, listTitle: 'Demo List', webUrl: webUrl, uniqueId: expectedUniqueId }) });
    assert(loggerLogSpy.calledWith({
      Attachments: false,
      AuthorId: 3,
      ContentTypeId: '0x0100B21BD271A810EE488B570BE49963EA34',
      Created: '2018-03-15T10:43:10Z',
      EditorId: 3,
      GUID: 'ea093c7b-8ae6-4400-8b75-e2d01154dffc',
      Modified: '2018-03-15T10:43:10Z',
      Title: expectedTitle
    }));
  });

  it('returns listItemInstance object when list item is requested and with permissions', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, listTitle: 'Demo List', webUrl: webUrl, id: expectedId, withPermissions: true }) });
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item is requested with an output type of json, and a list of fields are specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    await command.action(logger, { options: commandOptionsSchema.parse({ listTitle: 'Demo List', webUrl: webUrl, id: expectedId, output: 'json', fields: 'ID,Modified' }) });
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item is requested with an output type of json, a list of fields with lookup field are specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if ((opts.url as string).indexOf('&$expand=') > -1) {
        actualId = parseInt(opts.url.match(/\/items\((\d+)\)/i)[1]);
        return {
          "ID": actualId,
          "Modified": "2018-03-15T10:43:10Z",
          "Title": expectedTitle,
          "Company": `{ "Title": "Contoso" }`
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ listTitle: 'Demo List', webUrl: webUrl, id: expectedId, output: 'json', fields: 'Title,Modified,Company/Title' }) });
    assert.deepStrictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify({
      "Modified": "2018-03-15T10:43:10Z",
      "Title": expectedTitle,
      "Company": `{ "Title": "Contoso" }`
    }));
  });

  it('returns listItemInstance object when list item is requested with an output type of text, and no list of fields', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    await command.action(logger, { options: commandOptionsSchema.parse({ listTitle: 'Demo List', webUrl: webUrl, id: expectedId, output: 'text' }) });
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item is requested with an output type of text, and a list of fields specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    await command.action(logger, { options: commandOptionsSchema.parse({ listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', webUrl: webUrl, id: expectedId, output: 'json' }) });
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item is requested with an output type of text from a list specified by url, and a list of fields are being specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, webUrl: webUrl, id: expectedId, listUrl: listUrl, output: 'json', fields: 'Title,Modified' }) });
    assert.strictEqual(actualId, expectedId);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', webUrl: webUrl, id: expectedId, output: 'json' }) }),
      new CommandError('An error has occurred'));
  });
});
