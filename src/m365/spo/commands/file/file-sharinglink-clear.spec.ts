import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { GraphFileDetails } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command from './file-sharinglink-clear.js';
import spoFileSharingLinkListCommand from './file-sharinglink-list.js';

describe(commands.FILE_SHARINGLINK_CLEAR, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/demo';
  const fileUrl = '/sites/demo/Shared Documents/document.docx';
  const fileId = 'daebb04b-a773-4baa-b1d1-3625418e3234';
  const fileInformationResponse: GraphFileDetails = {
    SiteId: '9798e615-a586-455e-8486-84913f492c49',
    VroomDriveID: 'b!FeaYl4alXkWEhoSRP0ksSSOaj9osSfFPqj5bQNdluvlwfL79GNVISZZCf6nfB3vY',
    VroomItemID: '01A5WCPNXHFAS23ZNOF5D3XU2WU7S3I2AU'
  };
  const sharingLink = { "id": "8c2c9168-7d3d-4119-bcab-3c5340ce603b", "roles": ["read"], "hasPassword": false, "grantedToIdentitiesV2": [{ "group": { "displayName": "h Members", "email": "h@contoso.onmicrosoft.com", "id": "94da1e01-bbab-41e9-b9a4-4595c5805a6b" }, "siteUser": { "displayName": "h Members", "email": "h@contoso.onmicrosoft.com", "id": "428", "loginName": "c:0o.c|federateddirectoryclaimprovider|94da1e01-bbab-41e9-b9a4-4595c5805a6b" } }], "grantedToIdentities": [{ "user": { "displayName": "h Members", "email": "h@contoso.onmicrosoft.com", "id": "94da1e01-bbab-41e9-b9a4-4595c5805a6b" } }], "link": { "scope": "anonymous", "type": "view", "webUrl": "https://contoso.sharepoint.com/:b:/s/pnpcoresdktestgroup/EY50lub3559MtRKfj2hrZqoBea_L-lv1lND19RSCJGtWNg", "preventsDownload": false } };
  const sharingLinksListCommandResponse = {
    "stdout": JSON.stringify([sharingLink]),
    "stderr": ""
  };

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

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
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.get,
      Cli.promptForConfirmation,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
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
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl } });

    assert(postSpy.notCalled);
  });

  it('prompts before clearing the sharing links to a file when confirm option not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId } });

    assert(promptIssued);
  });

  it('clears sharing links from a specific file retrieved by url', async () => {
    const fileServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

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
      throw 'Error occurred while executing the command.';
    });

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${fileInformationResponse.SiteId}/drives/${fileInformationResponse.VroomDriveID}/items/${fileInformationResponse.VroomItemID}/permissions/${sharingLink.id}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: webUrl, fileUrl: fileUrl, force: true } });
    assert(deleteStub.called);
  });

  it('clears sharing links of type anonymous from a specific file retrieved by id', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

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
      throw 'Error occurred while executing the command.';
    });

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${fileInformationResponse.SiteId}/drives/${fileInformationResponse.VroomDriveID}/items/${fileInformationResponse.VroomItemID}/permissions/${sharingLink.id}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: webUrl, fileId: fileId, scope: 'anonymous' } });
    assert(deleteStub.called);
  });

  it('throws error when file not found by id', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        throw { error: { 'odata.error': { message: { value: 'File Not Found.' } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, fileId: fileId, verbose: true } } as any), new CommandError(`File Not Found.`));
  });
});
