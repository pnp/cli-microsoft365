import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
import commands from '../../commands.js';
import command from './file-sharinglink-add.js';

describe(commands.FILE_SHARINGLINK_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com';
  const fileId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';
  const fileUrl = '/sites/project-x/documents/SharedFile.docx';
  const fileDetailsResponse: GraphFileDetails = {
    SiteId: "0f9b8f4f-0e8e-4630-bb0a-501442db9b64",
    VroomItemID: "013TMHP6UOOSLON57HT5GLKEU7R5UGWZVK",
    VroomDriveID: "b!T4-bD44OMEa7ClAUQtubZID9tc40pGJKpguycvELod_Gx-lo4ZQiRJ7vylonTufG"
  };
  const graphResponseOne = {
    "id": "2a021f54-90a2-4016-b3b3-5f34d2e7d932",
    "roles": [
      "read"
    ],
    "hasPassword": false,
    "grantedToIdentitiesV2": [],
    "grantedToIdentities": [],
    "link": {
      "scope": "anonymous",
      "type": "view",
      "webUrl": "https://contoso.sharepoint.com/:b:/s/pnpcoresdktestgroup/EY50lub3559MtRKfj2hrZqoBWnHOpGIcgi4gzw9XiWYJ-A",
      "preventsDownload": false
    }
  };

  const graphResponseTwo = {
    "id": "2a021f54-90a2-4016-b3b3-5f34d2e7d932",
    "roles": [
      "read"
    ],
    "hasPassword": false,
    "grantedToIdentitiesV2": [],
    "grantedToIdentities": [],
    "link": {
      "scope": "organization",
      "type": "edit",
      "webUrl": "https://contoso.sharepoint.com/:b:/s/pnpcoresdktestgroup/EY50lub3559MtRKfj2hrZqoBWnHOpGIcgi4gzw9XiWYJ-A",
      "preventsDownload": false
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_SHARINGLINK_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a sharing link from a file specified by the id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${fileDetailsResponse.SiteId}/drives/${fileDetailsResponse.VroomDriveID}/items/${fileDetailsResponse.VroomItemID}/createLink`) {
        return graphResponseOne;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId, type: 'view', scope: 'anonymous', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponseOne));
  });

  it('creates a sharing link from a file specified by the url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${fileDetailsResponse.SiteId}/drives/${fileDetailsResponse.VroomDriveID}/items/${fileDetailsResponse.VroomItemID}/createLink`) {
        return graphResponseTwo;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, type: 'edit', expirationDateTime: "2023-01-09T16:20:00Z", scope: 'anonymous', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponseTwo));
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

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, type: 'view' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: 'invalid', type: 'view' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the expirationDateTime option is not a valid date', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, expirationDateTime: 'invalid date', type: 'view' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid scope specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, scope: 'invalid scope', type: 'view' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid type specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, type: 'invalid type' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if expirationDateTime option is specified for a non anonymous link', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, expirationDateTime: '2023-01-09T16:20:00Z', type: 'view', scope: 'organization' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the expirationDateTime option a valid date along with the anonymous scope', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, expirationDateTime: '2023-01-09T16:20:00Z', type: 'view', scope: 'anonymous' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, type: 'view' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
