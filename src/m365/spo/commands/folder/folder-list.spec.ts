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
import { urlUtil } from '../../../../utils/urlUtil';
import commands from '../../commands';
const command: Command = require('./folder-list');

describe(commands.FOLDER_LIST, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const parentFolderUrl = '/Shared Documents';
  const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, parentFolderUrl);
  const requestUrl: string = `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(serverRelativeUrl)}'&$skip=0&$top=5000`;

  const folderListOutput = {
    value: [
      { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 2, "Name": "Test", "ProgID": null, "ServerRelativeUrl": "/sites/abc/Shared Documents/Test", "TimeCreated": "2018-04-23T21:29:40Z", "TimeLastModified": "2018-04-23T21:32:13Z", "UniqueId": "3e735407-9c9f-418b-8378-450a9888d815", "WelcomePage": "" },
      { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "john", "ProgID": null, "ServerRelativeUrl": "/sites/abc/Shared Documents/john", "TimeCreated": "2018-05-02T22:28:50Z", "TimeLastModified": "2018-05-02T22:36:14Z", "UniqueId": "edeb37c6-8502-4a35-9fa2-6934bfc30214", "WelcomePage": "" },
      { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test111", "ProgID": null, "ServerRelativeUrl": "/sites/abc/Shared Documents/test111", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" },
      { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "Forms", "ProgID": null, "ServerRelativeUrl": "/sites/abc/Shared Documents/Forms", "TimeCreated": "2018-02-15T13:57:52Z", "TimeLastModified": "2018-02-15T13:57:52Z", "UniqueId": "cbb96da6-c2d8-4af0-9451-d534d5949371", "WelcomePage": "" }
    ]
  };

  const folderListOutputSingleFolder = {
    value: [
      { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 2, "Name": "Test", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/Test", "TimeCreated": "2018-04-23T21:29:40Z", "TimeLastModified": "2018-04-23T21:32:13Z", "UniqueId": "3e735407-9c9f-418b-8378-450a9888d815", "WelcomePage": "" }
    ]
  };

  const folderListOutputRecursiveLevel1 = {
    value: [
      { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 2, "Name": "Test2", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/Test/Test2", "TimeCreated": "2018-04-23T21:29:40Z", "TimeLastModified": "2018-04-23T21:32:13Z", "UniqueId": "3e735407-9c9f-418b-8378-450a9888d815", "WelcomePage": "" },
      { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "Test3", "ProgID": null, "ServerRelativeUrl": "/Shared Documents/Test/Test3", "TimeCreated": "2018-05-02T22:28:50Z", "TimeLastModified": "2018-05-02T22:36:14Z", "UniqueId": "edeb37c6-8502-4a35-9fa2-6934bfc30214", "WelcomePage": "" }
    ]
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Name', 'ServerRelativeUrl']);
  });

  it('should correctly handle folder get reject request', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === requestUrl) {
        throw 'error1';
      }
      throw `Invalid request ${opts.url}`;
    });
    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        parentFolderUrl: parentFolderUrl
      }
    } as any), new CommandError('error1'));
  });

  it('should correctly handle folder get success request', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === requestUrl) {
        return folderListOutput;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        parentFolderUrl: parentFolderUrl
      }
    });
    assert(loggerLogSpy.calledWith(folderListOutput.value));
  });

  it('retrieves folders with filter and fields option, requesting the ListItemAllFields Id property', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(serverRelativeUrl)}'&$skip=0&$top=5000&$expand=ListItemAllFields&$select=ListItemAllFields/Id,Name&$filter=name eq 'Folder1'`) {
        return {
          value: [
            {
              ListItemAllFields: {
                Id: 1,
                ID: 1
              },
              Name: "Test1"
            },
            {
              ListItemAllFields: {
                Id: 2
              },
              Name: "Test2"
            },
            {
              Name: "Test3"
            }
          ]
        };
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        parentFolderUrl: parentFolderUrl,
        filter: `name eq 'Folder1'`,
        fields: 'ListItemAllFields/Id,Name'
      }
    });
    assert(loggerLogSpy.calledWith([{ ListItemAllFields: { Id: 1 }, Name: "Test1" }, { ListItemAllFields: { Id: 2 }, Name: "Test2" }, { Name: "Test3" }]));
  });

  it('should correctly handle folder get success request with threshold limit', async () => {
    const folderThresholdLimitOutput = {
      value: new Array(5000).fill(folderListOutputSingleFolder.value)
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(serverRelativeUrl)}'&$skip=0&$top=5000`) {
        return folderThresholdLimitOutput;
      }

      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(serverRelativeUrl)}'&$skip=5000&$top=5000`) {
        return folderListOutput;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        parentFolderUrl: parentFolderUrl
      }
    });
    assert(loggerLogSpy.calledWith([...folderThresholdLimitOutput.value, ...folderListOutput.value]));
  });

  it('returns all information for output type json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === requestUrl) {
        return folderListOutput;
      }
      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        parentFolderUrl: parentFolderUrl,
        output: 'json'
      }
    });
    assert(loggerLogSpy.calledWith(folderListOutput.value));
  });

  it('returns all information recursive for output type json', async () => {
    const serverRelativeUrlLevel1First: string = `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(webUrl, `${parentFolderUrl}/Test`))}'&$skip=0&$top=5000`;
    const serverRelativeUrlLevel2First: string = `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(webUrl, `${parentFolderUrl}/Test/Test2`))}'&$skip=0&$top=5000`;
    const serverRelativeUrlLevel2Second: string = `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(webUrl, `${parentFolderUrl}/Test/Test3`))}'&$skip=0&$top=5000`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === requestUrl) {
        return folderListOutputSingleFolder;
      }

      if (opts.url === serverRelativeUrlLevel1First) {
        return folderListOutputRecursiveLevel1;
      }

      if (opts.url === serverRelativeUrlLevel2First) {
        return { value: [] };
      }

      if (opts.url === serverRelativeUrlLevel2Second) {
        return { value: [] };
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        parentFolderUrl: parentFolderUrl,
        recursive: true
      }
    });
    const expectedResults = folderListOutputSingleFolder.value;
    folderListOutputRecursiveLevel1.value.forEach(element => {
      expectedResults.push(element);
    });

    assert(loggerLogSpy.calledWith(expectedResults));
  });

  it('should send correct request params when /', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === requestUrl) {
        return folderListOutput;
      }
      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        parentFolderUrl: parentFolderUrl
      }
    });
    assert(loggerLogSpy.calledWith(folderListOutput.value));
  });

  it('should send correct request params when /sites/abc', async () => {
    const webUrl = 'https://contoso.sharepoint.com/sites/abc';
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, parentFolderUrl);
    const requestUrl: string = `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(serverRelativeUrl)}'&$skip=0&$top=5000`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === requestUrl) {
        return folderListOutput;
      }

      throw `Invalid request ${opts.url}`;
    });
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        parentFolderUrl: parentFolderUrl
      }
    });

    assert(loggerLogSpy.lastCall.calledWith(folderListOutput.value));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', parentFolderUrl: parentFolderUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and parentFolderUrl specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, parentFolderUrl: parentFolderUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
