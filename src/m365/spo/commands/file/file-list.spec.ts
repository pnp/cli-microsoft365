import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./file-list');

describe(commands.FILE_LIST, () => {
  const folder = 'Shared Documents';
  const fileResponse = {
    value: [
      {
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
      }
    ]
  };

  const fileThresholdResponse = {
    value: new Array(5000).fill({
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
    })
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

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves files from a folder when --recursive option is not supplied and output option is json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files?$skip=0&$top=5000`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: folder
      }
    });
    assert(loggerLogSpy.calledWith(fileResponse.value));
  });

  it('retrieves files from a folder when with filter and fields option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files?$skip=0&$top=5000&$expand=ListItemAllFields&$select=ListItemAllFields/Id&$filter=name eq 'Test.docx'`) {
        return {
          value: [
            {
              Id: 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: folder,
        filter: `name eq 'Test.docx'`,
        fields: 'ListItemAllFields/Id'
      }
    });
    assert(loggerLogSpy.calledWith([{ Id: 1 }]));
  });

  it('retrieves files from a folder with the list item threshold', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files?$skip=0&$top=5000`) {
        return fileThresholdResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files?$skip=5000&$top=5000`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: folder
      }
    });
    assert(loggerLogSpy.calledWith([...fileThresholdResponse.value, ...fileResponse.value]));
  });

  it('retrieves files from a folder with the folder threshold', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files?$skip=0&$top=5000`) {
        return fileResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Folders?$skip=0&$top=5000`) {
        return {
          value: new Array(5000).fill(
            {
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
            }
          )
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Folders?$skip=5000&$top=5000`) {
        return {
          value: [
            {
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
            }
          ]
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Folders?$skip=0&$top=5000`) {
        return {
          value: []
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Files?$skip=0&$top=5000`) {
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
        folder: folder
      }
    });
    assert(loggerLogSpy.calledWith(fileResponse.value));
  });

  it('retrieves files from a folder when --recursive option is not supplied and output option is text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files?$skip=0&$top=5000&$select=UniqueId,Name,ServerRelativeUrl`) {
        return fileTextResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'text',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared Documents'
      }
    });
    assert(loggerLogSpy.calledWith(fileTextResponse.value));
  });

  // Test for --recursive option. Uses onCall() method on stub to simulate recursion
  it('retrieves files from a folder and all the folders below it recursively when --recursive option is supplied and output option is json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Folders?$skip=0&$top=5000`) {
        return {
          value: [
            {
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
            }
          ]
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Folders?$skip=0&$top=5000`) {
        return {
          value: []
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files?$skip=0&$top=5000`) {
        return fileResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Files?$skip=0&$top=5000`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared Documents',
        recursive: true
      }
    });
    assert(loggerLogSpy.calledWith([...fileResponse.value, ...fileResponse.value]));
  });

  it('retrieves files from a folder and all the folders below it recursively when --recursive option is supplied and output option is text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(`Shared Documents/Fo'lde'r`)}')/Folders?$skip=0&$top=5000`) {
        return {
          value: [
            {
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
            }
          ]
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Folders?$skip=0&$top=5000`) {
        return {
          value: []
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(`Shared Documents/Fo'lde'r`)}')/Files?$skip=0&$top=5000&$select=UniqueId,Name,ServerRelativeUrl`) {
        return fileTextResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%20documents%2FLevel1-Folder')/Files?$skip=0&$top=5000&$select=UniqueId,Name,ServerRelativeUrl`) {
        return fileTextResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'text',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: `Shared Documents/Fo'lde'r`,
        recursive: true
      }
    });
    assert(loggerLogSpy.calledWith([...fileTextResponse.value, ...fileTextResponse.value]));
  });

  it('command correctly handles files list reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(err));
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

  it('supports specifying recursive', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--recursive') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folder: '/' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folder: '/' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
