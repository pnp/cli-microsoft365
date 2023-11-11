import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './file-list.js';

describe(commands.FILE_LIST, () => {
  const folderUrl = 'Shared Documents';

  //#region Mocked Responses
  const fileMock = {
    CheckInComment: '',
    CheckOutType: 2,
    ContentTag: '{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12',
    CustomizedPageStatus: 0,
    ETag: '\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"',
    Exists: true,
    IrmEnabled: false,
    Length: '331673',
    Level: 1,
    LinkingUri: 'https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b',
    LinkingUrl: 'https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b',
    MajorVersion: 3,
    MinorVersion: 0,
    Name: 'Test.docx',
    ServerRelativeUrl: '/sites/project-x/Shared documents/Test.docx',
    TimeCreated: '2018-02-05T08:42:36Z',
    TimeLastModified: '2018-02-05T08:44:03Z',
    Title: '',
    UIVersion: 1536,
    UIVersionLabel: '3.0',
    UniqueId: 'f09c4efe-b8c0-4e89-a166-03418661b89b'
  };

  const folderMock = {
    Exists: true,
    IsWOPIEnabled: false,
    ItemCount: 2,
    Name: 'Level1-Folder',
    ProgID: null,
    ServerRelativeUrl: '/sites/project-x/Shared documents/Level1-Folder',
    TimeCreated: '2021-05-22T09:00:33Z',
    TimeLastModified: '2021-05-24T09:08:33Z',
    UniqueId: 'cb9153af-b2f4-4d03-8798-020e98a3676d',
    WelcomePage: ''
  };

  const fileShortArrayResponse = {
    value: [fileMock]
  };

  const fileFullPageResponse = {
    value: new Array(5000).fill(fileMock)
  };

  const folderShortArrayResponse = {
    value: [folderMock]
  };

  const folderFullPageResponse = {
    value: new Array(5000).fill(folderMock)
  };

  const fileTextResponse = {
    value: [
      {
        UniqueId: 'f09c4efe-b8c0-4e89-a166-03418661b89b',
        Name: 'Test.docx',
        ServerRelativeUrl: '/sites/project-x/Shared documents/Test.docx'
      }
    ]
  };
  //#endregion

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves files from a folder when --recursive option is not supplied', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=0&$top=5000`) {
        return fileShortArrayResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folderUrl: folderUrl
      }
    });
    assert(loggerLogSpy.calledWith(fileShortArrayResponse.value));
  });

  it('retrieves empty list of files from a folder', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=0&$top=5000`) {
        return { value: [] };
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folderUrl: folderUrl
      }
    });
    assert(loggerLogSpy.calledWith([]));
  });

  it('retrieves files from a folder with filter and fields option, requesting the ListItemAllFields Id property', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=0&$top=5000&$expand=ListItemAllFields&$select=ListItemAllFields/Id,Name&$filter=name eq 'Test.docx'`) {
        return {
          value: [
            {
              ListItemAllFields: {
                Id: 1,
                ID: 1
              },
              Name: "Test.docx"
            },
            {
              ListItemAllFields: {
                Id: 2
              },
              Name: "Test.docx"
            },
            {
              Name: "Test.docx"
            }
          ]
        };
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folderUrl: folderUrl,
        filter: `name eq 'Test.docx'`,
        fields: 'ListItemAllFields/Id,Name'
      }
    });
    assert(loggerLogSpy.calledWith([{ ListItemAllFields: { Id: 1 }, Name: "Test.docx" }, { ListItemAllFields: { Id: 2 }, Name: "Test.docx" }, { Name: "Test.docx" }]));
  });

  it('retrieves files from a folder with filter and fields option, requesting the ListItemAllFields Title property', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=0&$top=5000&$expand=ListItemAllFields&$select=ListItemAllFields/Title&$filter=name eq 'Test.docx'`) {
        return {
          value: [
            {
              ListItemAllFields: {
                Title: 'Test title'
              }
            }
          ]
        };
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folderUrl: folderUrl,
        filter: `name eq 'Test.docx'`,
        fields: 'ListItemAllFields/Title'
      }
    });
    assert(loggerLogSpy.calledWith([{ ListItemAllFields: { Title: 'Test title' } }]));
  });

  it('retrieves files from a folder in multiple pages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=0&$top=5000`) {
        return fileFullPageResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=5000&$top=5000`) {
        return fileShortArrayResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folderUrl: folderUrl
      }
    });
    assert(loggerLogSpy.calledWith([...fileFullPageResponse.value, ...fileShortArrayResponse.value]));
  });

  it('retrieves files from a folder recursively in multiple pages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=0&$top=5000`) {
        return fileShortArrayResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Folders?$skip=0&$top=5000&$select=ServerRelativeUrl`) {
        return folderFullPageResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Folders?$skip=5000&$top=5000&$select=ServerRelativeUrl`) {
        return folderShortArrayResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Folders?$skip=0&$top=5000&$select=ServerRelativeUrl`) {
        return {
          value: []
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Files?$skip=0&$top=5000`) {
        return {
          value: []
        };
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        recursive: true,
        folderUrl: folderUrl
      }
    });
    assert(loggerLogSpy.calledWith(fileShortArrayResponse.value));
  });

  it('retrieves files from a folder when --recursive option is not supplied and output option is text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=0&$top=5000&$select=UniqueId,Name,ServerRelativeUrl`) {
        return fileTextResponse;
      }
      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        output: 'text',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folderUrl: 'Shared Documents'
      }
    });
    assert(loggerLogSpy.calledWith(fileTextResponse.value));
  });

  // Test for --recursive option. Uses onCall() method on stub to simulate recursion
  it('retrieves files from a folder and all the folders below it recursively when --recursive option is supplied and output option is json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Folders?$skip=0&$top=5000&$select=ServerRelativeUrl`) {
        return folderShortArrayResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Folders?$skip=0&$top=5000&$select=ServerRelativeUrl`) {
        return {
          value: []
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/project-x/' + folderUrl)}')/Files?$skip=0&$top=5000`) {
        return fileShortArrayResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Files?$skip=0&$top=5000`) {
        return fileShortArrayResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folderUrl: 'Shared Documents',
        recursive: true
      }
    });
    assert(loggerLogSpy.calledWith([...fileShortArrayResponse.value, ...fileShortArrayResponse.value]));
  });

  it('retrieves files from a folder and all the folders below it recursively when --recursive option is supplied and output option is text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(`/sites/project-x/Shared Documents`)}')/Folders?$skip=0&$top=5000&$select=ServerRelativeUrl`) {
        return folderShortArrayResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Folders?$skip=0&$top=5000&$select=ServerRelativeUrl`) {
        return {
          value: []
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(`/sites/project-x/Shared Documents`)}')/Files?$skip=0&$top=5000&$select=UniqueId,Name,ServerRelativeUrl`) {
        return fileTextResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Files?$skip=0&$top=5000&$select=UniqueId,Name,ServerRelativeUrl`) {
        return fileTextResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        output: 'text',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folderUrl: 'Shared Documents',
        recursive: true
      }
    });
    assert(loggerLogSpy.calledWith([...fileTextResponse.value, ...fileTextResponse.value]));
  });

  it('command correctly handles files list reject request', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativePath') > -1) {
        throw error;
      }

      throw `Invalid request ${opts.url}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: `Shared Documents/Folder`
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: '/' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});