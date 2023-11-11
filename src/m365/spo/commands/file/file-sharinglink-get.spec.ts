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
import command from './file-sharinglink-get.js';

describe(commands.FILE_SHARINGLINK_GET, () => {
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
