import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import { formatting } from '../../../../utils/formatting';
import { GraphFileDetails } from './GraphFileDetails';
import { urlUtil } from '../../../../utils/urlUtil';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as spoFileSharingLinkListCommand from './file-sharinglink-list';
const command: Command = require('./file-sharinglink-clear');

describe(commands.FILE_SHARINGLINK_CLEAR, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/demo';
  const fileUrl = '/sites/demo/Shared Documents/document.docx';
  const fileId = 'daebb04b-a773-4baa-b1d1-3625418e3234';
  const fileInformationResponse: GraphFileDetails = {
    SiteId: '9798e615-a586-455e-8486-84913f492c49',
    VroomDriveID: 'b!FeaYl4alXkWEhoSRP0ksSSOaj9osSfFPqj5bQNdluvlwfL79GNVISZZCf6nfB3vY',
    VroomItemID: '01A5WCPNXHFAS23ZNOF5D3XU2WU7S3I2AU'
  };
  const sharingLink = { "id": "8c2c9168-7d3d-4119-bcab-3c5340ce603b", "roles": ["read"], "hasPassword": false, "grantedToIdentitiesV2": [{ "group": { "displayName": "h Members", "email": "h@mathijsdev2.onmicrosoft.com", "id": "94da1e01-bbab-41e9-b9a4-4595c5805a6b" }, "siteUser": { "displayName": "h Members", "email": "h@mathijsdev2.onmicrosoft.com", "id": "428", "loginName": "c:0o.c|federateddirectoryclaimprovider|94da1e01-bbab-41e9-b9a4-4595c5805a6b" } }], "grantedToIdentities": [{ "user": { "displayName": "h Members", "email": "h@mathijsdev2.onmicrosoft.com", "id": "94da1e01-bbab-41e9-b9a4-4595c5805a6b" } }], "link": { "scope": "anonymous", "type": "view", "webUrl": "https://mathijsdev2.sharepoint.com/:b:/s/pnpcoresdktestgroup/EY50lub3559MtRKfj2hrZqoBea_L-lv1lND19RSCJGtWNg", "preventsDownload": false } };
  const sharingLinksListCommandResponse = {
    "stdout": JSON.stringify([sharingLink]),
    "stderr": ""
  };

  let log: any[];
  let logger: Logger;
  //let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;

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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
    //loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      Cli.prompt,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name, commands.FILE_SHARINGLINK_CLEAR);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl and fileId options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the scope option is not a valid scope', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, scope: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the scope option is a valid scope', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, scope: 'users' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('aborts clearing the sharing links to a file when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl } });

    assert(postSpy.notCalled);
  });

  it('prompts before clearing the sharing links to a file when confirm option not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId } });

    let promptIssued = false;
    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('clears all sharing links from a specific file retrieved by url', async () => {
    const fileServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileServerRelativeUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileInformationResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFileSharingLinkListCommand) {
        return sharingLinksListCommandResponse;
      }
      throw 'Error occured while executing the command.';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: webUrl, fileUrl: fileUrl, confirm: true } });
    assert(postStub.called);
  });

  it('clears sharing links of type anonymous from a specific file retrieved by id', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileInformationResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFileSharingLinkListCommand) {
        return sharingLinksListCommandResponse;
      }
      throw 'Error occured while executing the command.';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: webUrl, fileId: fileId, scope: 'anonymous' } });
    assert(postStub.called);

  });

  it('clears sharing links in multiple batches if more than 20 sharing links are found', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileInformationResponse;
      }

      throw 'Invalid request';
    });

    const sharingLinkResponseLarge: any[] = [];
    for (let i = 0; i < 50; i++) {
      sharingLinkResponseLarge.push(sharingLink);
    }
    let amountOfBatches = 0;

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFileSharingLinkListCommand) {
        return { "stdout": JSON.stringify(sharingLinkResponseLarge) };
      }
      throw 'Error occured while executing the command.';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        amountOfBatches++;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: webUrl, fileId: fileId, scope: 'anonymous' } });
    assert.strictEqual(amountOfBatches, 3);
  });

  it('throws error when no sharing links are found', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileInformationResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoFileSharingLinkListCommand) {
        return { "stdout": "[]" };
      }
      throw 'Error occured while executing the command.';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, fileId: fileId, verbose: true } } as any), new CommandError(`There are no sharing links to be removed for the specific file.`));
  });

  it('throws error when file not found by id', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        throw { error: { 'odata.error': { message: { value: 'File Not Found.' } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, fileId: fileId, verbose: true } } as any), new CommandError(`File Not Found.`));
  });
});
