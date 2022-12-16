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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./file-list');

describe(commands.FILE_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves files from a folder when --recursive option is not supplied and output option is json', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {
            "Files": [{
              "CheckInComment": "",
              "CheckOutType": 2,
              "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
              "CustomizedPageStatus": 0,
              "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
              "Exists": true,
              "IrmEnabled": false,
              "Length": "331673",
              "Level": 1,
              "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
              "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
              "MajorVersion": 3,
              "MinorVersion": 0,
              "Name": "Test.docx",
              "ServerRelativeUrl": "/sites/project-x/Shared documents/Test.docx",
              "TimeCreated": "2018-02-05T08:42:36Z",
              "TimeLastModified": "2018-02-05T08:44:03Z",
              "Title": "",
              "UIVersion": 1536,
              "UIVersionLabel": "3.0",
              "UniqueId": "f09c4efe-b8c0-4e89-a166-03418661b89b"
            }],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Shared Documents",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared Documents'
      }
    });
    assert(loggerLogSpy.calledWith([{
      CheckInComment: "",
      CheckOutType: 2,
      ContentTag: "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
      CustomizedPageStatus: 0,
      ETag: "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
      Exists: true,
      IrmEnabled: false,
      Length: "331673",
      Level: 1,
      LinkingUri: "https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
      LinkingUrl: "https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
      MajorVersion: 3,
      MinorVersion: 0,
      Name: "Test.docx",
      ServerRelativeUrl: "/sites/project-x/Shared documents/Test.docx",
      TimeCreated: "2018-02-05T08:42:36Z",
      TimeLastModified: "2018-02-05T08:44:03Z",
      Title: "",
      UIVersion: 1536,
      UIVersionLabel: "3.0",
      UniqueId: "f09c4efe-b8c0-4e89-a166-03418661b89b"
    }]));
  });

  it('retrieves files from a folder when --recursive option is not supplied and output option is text', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {
            "Files": [
              {
                "UniqueId": "f09c4efe-b8c0-4e89-a166-03418661b89b",
                "Name": "Test.docx",
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Test.docx"
              }
            ],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Shared Documents",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        output: 'text',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared Documents'
      }
    });
    assert(loggerLogSpy.calledWith(
      [{
        UniqueId: 'f09c4efe-b8c0-4e89-a166-03418661b89b',
        Name: 'Test.docx',
        ServerRelativeUrl: '/sites/project-x/Shared documents/Test.docx'
      }]
    ));
  });

  // Test for --recursive option. Uses onCall() method on stub to simulate recursion
  it('retrieves files from a folder and all the folders below it recursively when --recursive option is supplied and output option is json', async () => {

    const requestStub = sinon.stub(request, 'get');

    // Represents the first call which returns files and a folder
    requestStub.onCall(0).callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {
            "Files": [{
              "CheckInComment": "",
              "CheckOutType": 2,
              "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
              "CustomizedPageStatus": 0,
              "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
              "Exists": true,
              "IrmEnabled": false,
              "Length": "331673",
              "Level": 1,
              "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
              "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
              "MajorVersion": 3,
              "MinorVersion": 0,
              "Name": "Test.docx",
              "ServerRelativeUrl": "/sites/project-x/Shared documents/Test.docx",
              "TimeCreated": "2018-02-05T08:42:36Z",
              "TimeLastModified": "2018-02-05T08:44:03Z",
              "Title": "",
              "UIVersion": 1536,
              "UIVersionLabel": "3.0",
              "UniqueId": "f09c4efe-b8c0-4e89-a166-03418661b89b"
            }],
            "Folders": [
              {
                "Exists": true,
                "IsWOPIEnabled": false,
                "ItemCount": 2,
                "Name": "Level1-Folder",
                "ProgID": null,
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder",
                "TimeCreated": "2021-05-22T09:00:33Z",
                "TimeLastModified": "2021-05-24T09:08:33Z",
                "UniqueId": "cb9153af-b2f4-4d03-8798-020e98a3676d",
                "WelcomePage": ""
              }
            ],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Shared Documents",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    // Represents the second call which returns only files
    requestStub.onCall(1).callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {
            "Files": [
              {
                "CheckInComment": "",
                "CheckOutType": 2,
                "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
                "CustomizedPageStatus": 0,
                "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
                "Exists": true,
                "IrmEnabled": false,
                "Length": "331673",
                "Level": 1,
                "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
                "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
                "MajorVersion": 3,
                "MinorVersion": 0,
                "Name": "Test.docx",
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder/Level1-Test.docx",
                "TimeCreated": "2018-02-05T08:42:36Z",
                "TimeLastModified": "2018-02-05T08:44:03Z",
                "Title": "",
                "UIVersion": 1536,
                "UIVersionLabel": "3.0",
                "UniqueId": "1d0cae03-5ea7-438d-b4ad-3cbd62d52e46"
              }
            ],
            "Folders": [],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Shared Documents",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents/Level1-Folder",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared Documents',
        recursive: true
      }
    });
    assert(loggerLogSpy.calledWith(
      [{
        "CheckInComment": "",
        "CheckOutType": 2,
        "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
        "CustomizedPageStatus": 0,
        "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
        "Exists": true,
        "IrmEnabled": false,
        "Length": "331673",
        "Level": 1,
        "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
        "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
        "MajorVersion": 3,
        "MinorVersion": 0,
        "Name": "Test.docx",
        "ServerRelativeUrl": "/sites/project-x/Shared documents/Test.docx",
        "TimeCreated": "2018-02-05T08:42:36Z",
        "TimeLastModified": "2018-02-05T08:44:03Z",
        "Title": "",
        "UIVersion": 1536,
        "UIVersionLabel": "3.0",
        "UniqueId": "f09c4efe-b8c0-4e89-a166-03418661b89b"
      },
      {
        "CheckInComment": "",
        "CheckOutType": 2,
        "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
        "CustomizedPageStatus": 0,
        "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
        "Exists": true,
        "IrmEnabled": false,
        "Length": "331673",
        "Level": 1,
        "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
        "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
        "MajorVersion": 3,
        "MinorVersion": 0,
        "Name": "Test.docx",
        "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder/Level1-Test.docx",
        "TimeCreated": "2018-02-05T08:42:36Z",
        "TimeLastModified": "2018-02-05T08:44:03Z",
        "Title": "",
        "UIVersion": 1536,
        "UIVersionLabel": "3.0",
        "UniqueId": "1d0cae03-5ea7-438d-b4ad-3cbd62d52e46"
      }]
    ));
  });

  // Test for --recursive option. Uses onCall() method on stub to simulate recursion
  it('retrieves files from a folder and all the folders below it recursively when --recursive option is supplied and output option is text', async () => {
    const requestStub = sinon.stub(request, 'get');

    // Represents the first call which returns files and a folder
    requestStub.onCall(0).callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {
            "Files": [
              {
                "UniqueId": "f09c4efe-b8c0-4e89-a166-03418661b89b",
                "Name": "Test.docx",
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Test.docx"
              }
            ],
            "Folders": [
              {
                "Exists": true,
                "IsWOPIEnabled": false,
                "ItemCount": 2,
                "Name": "Level1-Folder",
                "ProgID": null,
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder",
                "TimeCreated": "2021-05-22T09:00:33Z",
                "TimeLastModified": "2021-05-24T09:08:33Z",
                "UniqueId": "cb9153af-b2f4-4d03-8798-020e98a3676d",
                "WelcomePage": ""
              }
            ],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Shared Documents",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    // Represents the second call which returns a second level deep folder
    requestStub.onCall(1).callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {
            "Files": [
              {
                "UniqueId": "1d0cae03-5ea7-438d-b4ad-3cbd62d52e46",
                "Name": "Level1-Test.docx",
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder/Level1-Test.docx"
              }
            ],
            "Folders": [
              {
                "Exists": true,
                "IsWOPIEnabled": false,
                "ItemCount": 2,
                "Name": "Level2-Folder",
                "ProgID": null,
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder/Level2-Folder",
                "TimeCreated": "2021-05-22T09:00:33Z",
                "TimeLastModified": "2021-05-24T09:08:33Z",
                "UniqueId": "cb9153af-b2f4-4d03-8798-020e98a3676d",
                "WelcomePage": ""
              }
            ],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Level1-Folder",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents/Level1-Folder",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    // Represents the third call which only retrieves files
    requestStub.onCall(2).callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {
            "Files": [
              {
                "UniqueId": "f65deb00-4d0e-44cc-a9db-027d54039b4d",
                "Name": "Level2-Test.docx",
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder/Level2-Folder/Level2-Test.docx"
              }
            ],
            "Folders": [],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Level2-Folder",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents/Level1-Folder/Level2-Folder",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        output: 'text',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared Documents',
        recursive: true
      }
    });
    assert(loggerLogSpy.calledWith(
      [{
        UniqueId: 'f09c4efe-b8c0-4e89-a166-03418661b89b',
        Name: 'Test.docx',
        ServerRelativeUrl: '/sites/project-x/Shared documents/Test.docx'
      },
      {
        UniqueId: '1d0cae03-5ea7-438d-b4ad-3cbd62d52e46',
        Name: 'Level1-Test.docx',
        ServerRelativeUrl: '/sites/project-x/Shared documents/Level1-Folder/Level1-Test.docx'
      },
      {
        "UniqueId": "f65deb00-4d0e-44cc-a9db-027d54039b4d",
        "Name": "Level2-Test.docx",
        "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder/Level2-Folder/Level2-Test.docx"
      }]
    ));
  });

  it('properly escapes single quotes in folder name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFolderByServerRelativeUrl('Shared%20Documents%2FFo''lde''r')`) > -1) {
        return Promise.resolve(
          {
            "Files": [
              {
                "UniqueId": "f09c4efe-b8c0-4e89-a166-03418661b89b",
                "Name": "Test.docx",
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Test.docx"
              }
            ],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Shared Documents",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        output: 'text',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: `Shared Documents/Fo'lde'r`
      }
    });
    assert(loggerLogSpy.calledWith(
      [{
        UniqueId: 'f09c4efe-b8c0-4e89-a166-03418661b89b',
        Name: 'Test.docx',
        ServerRelativeUrl: '/sites/project-x/Shared documents/Test.docx'
      }]
    ));
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

  it('uses correct API url when output json option is passed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('select123=') > -1) {
        return Promise.resolve('Correct Url1');
      }

      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {
            "Files": [
              {
                "UniqueId": "f65deb00-4d0e-44cc-a9db-027d54039b4d",
                "Name": "Level2-Test.docx",
                "ServerRelativeUrl": "/sites/project-x/Shared documents/Level1-Folder/Level2-Folder/Level2-Test.docx"
              }
            ],
            "Folders": [],
            "Exists": true,
            "IsWOPIEnabled": false,
            "ItemCount": 3,
            "Name": "Level2-Folder",
            "ProgID": null,
            "ServerRelativeUrl": "/sites/project-x/Shared Documents/Level1-Folder/Level2-Folder",
            "TimeCreated": "2021-05-22T08:58:37Z",
            "TimeLastModified": "2021-05-22T09:00:33Z",
            "UniqueId": "dee34261-95f0-49c0-9090-f8d2d581787c",
            "WelcomePage": ""
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'Shared Documents'
      }
    });
    assert('Correct Url');
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
