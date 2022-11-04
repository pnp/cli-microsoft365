import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { urlUtil } from '../../../../utils/urlUtil';
import commands from '../../commands';
const command: Command = require('./listitem-get');

describe(commands.LISTITEM_GET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listUrl = 'sites/project-x/documents';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const expectedTitle = `List Item 1`;
  const expectedId = 147;

  let actualId = 0;

  const getFakes = async (opts: any) => {
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
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_GET), true);
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

  it('fails validation if listTitle, listId or listUrl option not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: expectedId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle, listId and listUrl are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listUrl: listUrl, id: expectedId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listTitle: 'Demo List', id: expectedId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: expectedId } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo', id: expectedId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: expectedId } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the specified id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('returns listItemInstance object when list item is requested', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    const options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: webUrl,
      id: expectedId
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item is requested with an output type of json, and a list of fields are specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    const options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: webUrl,
      id: expectedId,
      output: "json",
      properties: "ID,Modified"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item is requested with an output type of json, a list of fields with lookup field are specified', async () => {
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if ((opts.url as string).indexOf('&$expand=') > -1) {
        actualId = parseInt(opts.url.match(/\/items\((\d+)\)/i)[1]);
        return Promise.resolve(
          {
            "ID": actualId,
            "Modified": "2018-03-15T10:43:10Z",
            "Title": expectedTitle,
            "Company": `{ "Title": "Contoso" }`
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    command.allowUnknownOptions();

    const options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: webUrl,
      id: expectedId,
      output: "json",
      properties: "Title,Modified,Company/Title"
    };

    await command.action(logger, { options: options } as any);
    assert.deepStrictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify({
      "Modified": "2018-03-15T10:43:10Z",
      "Title": expectedTitle,
      "Company": `{ "Title": "Contoso" }`
    }));
  });

  it('returns listItemInstance object when list item is requested with an output type of text, and no list of fields', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    const options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: webUrl,
      id: expectedId,
      output: "text"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item is requested with an output type of text, and a list of fields specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    const options: any = {
      debug: false,
      listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
      webUrl: webUrl,
      id: expectedId,
      output: "json"
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('returns listItemInstance object when list item is requested with an output type of text from a list specified by url, and a list of fields are being specified', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    const options: any = {
      verbose: true,
      webUrl: webUrl,
      id: expectedId,
      listUrl: listUrl,
      output: 'json',
      properties: 'Title,Modified'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(actualId, expectedId);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    const options: any = {
      debug: false,
      listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
      webUrl: webUrl,
      id: expectedId,
      output: "json"
    };

    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('An error has occurred'));
  });
});
