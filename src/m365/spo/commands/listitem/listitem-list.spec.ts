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
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command, { options } from './listitem-list.js';

describe(commands.LISTITEM_LIST, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listUrl = 'sites/project-x/documents';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const listItemResponse = {
    value:
      [{
        "Attachments": false,
        "AuthorId": 3,
        "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
        "Created": "2018-08-15T13:43:12Z",
        "EditorId": 3,
        "GUID": "2b6bd9e0-3c43-4420-891e-20053e3c4664",
        "Id": 1,
        "ID": 1,
        "Modified": "2018-08-15T13:43:12Z",
        "Title": "Example item 1"
      },
      {
        "Attachments": false,
        "AuthorId": 3,
        "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
        "Created": "2018-08-15T13:44:10Z",
        "EditorId": 3,
        "GUID": "47c5fc61-afb7-4081-aa32-f4386b8a86ea",
        "Id": 2,
        "ID": 2,
        "Modified": "2018-08-15T13:44:10Z",
        "Title": "Example item 2"
      }]
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const expectedArrayLength = 2;
  let returnArrayLength = 0;

  const postFakes = async (opts: any): Promise<any> => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('/GetItems') > -1) {
        returnArrayLength = 2;
        return opts.data.query.ListItemCollectionPosition === undefined ? listItemResponse : { value: [] };
      }
    }
    returnArrayLength = 0;
    throw 'Invalid request';
  };

  const getFakes = async (opts: any): Promise<any> => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('/items') > -1 && (opts.url as string).indexOf('$top=6') > -1) {
        returnArrayLength = 0;
        return { value: [] };
      }
      if ((opts.url as string).indexOf('/items') > -1) {
        returnArrayLength = 2;
        return listItemResponse;
      }
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$top=5000&$select=Title%2CID`) {
      returnArrayLength = 2;
      return listItemResponse;
    }
    returnArrayLength = 0;
    throw 'Invalid request';
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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
    assert.strictEqual(actual.error?.issues[0].message, 'foo is not a valid GUID');
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if listId, listTitle and listUrl are not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listTitle are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listTitle: 'Demo List' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if camlQuery and fields are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', camlQuery: '<Query><ViewFields><FieldRef Name="Title" /><FieldRef Name="Id" /></ViewFields></Query>', fields: 'Title,Id' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if camlQuery and pageSize are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', camlQuery: '<Query><RowLimit>2</RowLimit></Query>', pageSize: 3 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if camlQuery and pageNumber are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', camlQuery: '<Query><RowLimit>2</RowLimit></Query>', pageNumber: 3 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if pageNumber is specified and pageSize is not', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', pageNumber: 3 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the specific pageSize is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', pageSize: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the specific pageNumber is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', pageSize: 3, pageNumber: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Unrecognized key: "unknownOption"');
  });

  it('returns array of listItemInstance objects when a list of items is requested, and debug mode enabled', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a list of fields and a filter specified', async () => {
    const listTitle = `Test'list`;
    const filter = `Title eq 'Demo list item'`;
    const fields = 'Title,ID';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/items?$top=2&$filter=${encodeURIComponent(filter)}&$select=${formatting.encodeQueryParameter(fields)}`) {
        returnArrayLength = 2;
        return listItemResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: listTitle,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        output: 'json',
        pageSize: 2,
        filter: filter,
        fields: 'Title,ID'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
    returnArrayLength = 0;
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, a page number specified, a list of fields and a filter specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        output: 'json',
        pageSize: 2,
        pageNumber: 2,
        filter: "Title eq 'Demo list item",
        fields: 'Title,ID'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns empty array of listItemInstance objects when a list of items is requested with an output type of json, a page number specified, a list of fields and a filter specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        output: 'json',
        pageSize: 3,
        pageNumber: 2,
        filter: "Title eq 'Demo list item",
        fields: 'Title,ID'
      })
    });
    assert.strictEqual(returnArrayLength, 0);
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a pageNumber is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        output: 'json',
        pageSize: 2,
        pageNumber: 2,
        fields: 'Title,ID'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, and a list of fields specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        fields: 'Title,ID'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items by list url is requested with no output type specified, and a list of fields specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        listUrl: listUrl,
        webUrl: webUrl,
        fields: 'Title,ID'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, a list of fields with lookup field specified', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf('$expand=') > -1) {
        return Promise.resolve({
          value:
            [{
              "ID": 1,
              "Modified": "2018-08-15T13:43:12Z",
              "Title": "Example item 1",
              "Company": { "Title": "Contoso" }
            },
            {
              "ID": 2,
              "Modified": "2018-08-15T13:44:10Z",
              "Title": "Example item 2",
              "Company": { "Title": "Fabrikam" }
            }]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        fields: 'Title,Modified,Company/Title'
      })
    });
    assert.deepStrictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "Modified": "2018-08-15T13:43:12Z",
        "Title": "Example item 1",
        "Company": { "Title": "Contoso" }
      },
      {
        "Modified": "2018-08-15T13:44:10Z",
        "Title": "Example item 2",
        "Company": { "Title": "Fabrikam" }
      }
    ]));
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of text, and no fields specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        output: 'text'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and no fields specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        output: 'json'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with a camlQuery specified, and output set to json, and debug mode is enabled', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>",
        output: 'json'
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with a camlQuery specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
      })
    });
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
      })
    }),
    new CommandError('An error has occurred'));
  });
});
