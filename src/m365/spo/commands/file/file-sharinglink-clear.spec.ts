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
const command: Command = require('./file-sharinglink-get');

describe(commands.FILE_SHARINGLINK_CLEAR, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/demo';
  const fileUrl = '/sites/demo/Shared Documents/document.docx';
  const fileId = 'daebb04b-a773-4baa-b1d1-3625418e3234';
  const id = 'U1BEZW1vIFZpc2l0b3Jz';

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const fileInformationResponse: GraphFileDetails = {
    SiteId: '9798e615-a586-455e-8486-84913f492c49',
    VroomDriveID: 'b!FeaYl4alXkWEhoSRP0ksSSOaj9osSfFPqj5bQNdluvlwfL79GNVISZZCf6nfB3vY',
    VroomItemID: '01A5WCPNXHFAS23ZNOF5D3XU2WU7S3I2AU'
  };

  const fileSharingLinkResponse = {
    "id": id,
    "roles": [
      "read"
    ],
    "grantedToV2": {
      "siteGroup": {
        "displayName": "Demo Visitors",
        "id": "5",
        "loginName": "Demo Visitors"
      }
    },
    "grantedTo": {
      "user": {
        "displayName": "Demo Visitors"
      }
    },
    "inheritedFrom": {}
  };

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
    assert.strictEqual(command.name, commands.FILE_SHARINGLINK_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '12345', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('gets a specific sharing link of a file by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileInformationResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${fileInformationResponse.SiteId}/drives/${fileInformationResponse.VroomDriveID}/items/${fileInformationResponse.VroomItemID}/permissions/${id}`) {
        return fileSharingLinkResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        fileId: fileId,
        id: id
      }
    });

    assert(loggerLogSpy.calledWith(fileSharingLinkResponse));
  });

  it('gets a specific sharing link of a file by URL', async () => {
    const fileServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileServerRelativeUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileInformationResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${fileInformationResponse.SiteId}/drives/${fileInformationResponse.VroomDriveID}/items/${fileInformationResponse.VroomItemID}/permissions/${id}`) {
        return fileSharingLinkResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        id: id
      }
    });

    assert(loggerLogSpy.calledWith(fileSharingLinkResponse));
  });

  it('throws error when file not found by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        throw { error: { 'odata.error': { message: { value: 'File Not Found.' } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, fileId: fileId, verbose: true } } as any),
      new CommandError(`File Not Found.`));
  });
});
