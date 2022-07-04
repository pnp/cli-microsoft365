import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./listitem-list');

describe(commands.LISTITEM_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const expectedArrayLength = 2;
  let returnArrayLength = 0;

  const postFakes = (opts: any) => {
    if ((opts.url as string).indexOf('/GetItems') > -1) {
      returnArrayLength = 2;
      return Promise.resolve({
        value:
          [{
            "Attachments": false,
            "AuthorId": 3,
            "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
            "Created": "2018-08-15T13:43:12Z",
            "EditorId": 3,
            "GUID": "2b6bd9e0-3c43-4420-891e-20053e3c4664",
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
      });
    }
    returnArrayLength = 0;
    return Promise.reject('Invalid request');
  };

  const getFakes = (opts: any) => {
    if ((opts.url as string).indexOf('/items') > -1) {
      returnArrayLength = 2;
      return Promise.resolve({
        value:
          [{
            "Attachments": false,
            "AuthorId": 3,
            "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
            "Created": "2018-08-15T13:43:12Z",
            "EditorId": 3,
            "GUID": "2b6bd9e0-3c43-4420-891e-20053e3c4664",
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
            "ID": 2,
            "Id": 2,
            "Modified": "2018-08-15T13:44:10Z",
            "Title": "Example item 2"
          }]
      });
    }
    returnArrayLength = 0;
    return Promise.reject('Invalid request');
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
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

    loggerLogSpy = sinon.spy(logger, 'log');
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
      appInsights.trackEvent,
      pid.getProcessName,
      spo.getRequestDigest
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
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
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
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

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if camlQuery and fields are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', camlQuery: '<Query><ViewFields><FieldRef Name="Title" /><FieldRef Name="Id" /></ViewFields></Query>', fields: 'Title,Id' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if camlQuery and pageSize are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', camlQuery: '<Query><RowLimit>2</RowLimit></Query>', pageSize: 3 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if camlQuery and pageNumber are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', camlQuery: '<Query><RowLimit>2</RowLimit></Query>', pageNumber: 3 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if pageNumber is specified and pageSize is not', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', pageNumber: 3 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specific pageSize is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', pageSize: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specific pageNumber is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', pageSize: 3, pageNumber: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('returns array of listItemInstance objects when a list of items is requested, and debug mode enabled', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a list of fields and a filter specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      output: "json",
      pageSize: 2,
      filter: "Title eq 'Demo list item",
      fields: "Title,ID"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, a page number specified, a list of fields and a filter specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      output: "json",
      pageSize: 2,
      pageNumber: 2,
      filter: "Title eq 'Demo list item",
      fields: "Title,ID"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a pageNumber is specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      output: "json",
      pageSize: 2,
      pageNumber: 2,
      fields: "Title,ID"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, and a list of fields specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      fields: "Title,ID"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, a list of fields with lookup field specified', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf('&$expand=') > -1) {
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

    const options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      fields: "Title,Modified,Company/Title"
    };

    await command.action(logger, { options: options } as any);
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

    const options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      output: "text"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with a camlQuery specified, and output set to json, and debug mode is enabled', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>",
      output: "json"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with a camlQuery specified, and debug mode is disabled', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(returnArrayLength, expectedArrayLength);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    const options: any = {
      debug: false,
      listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('An error has occurred'));
  });
});
