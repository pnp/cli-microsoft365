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
import { GraphFileDetails } from '../../../../utils/spo';
import { odata } from '../../../../utils/odata';
const command: Command = require('./file-sharinglink-list');

describe(commands.FILE_SHARINGLINK_LIST, () => {
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
  const graphResponse = {
    value: [
      {
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
      },
      {
        "id": "a47e5387-8868-497c-bb00-115c66c60390",
        "roles": [
          "read"
        ],
        "hasPassword": true,
        "grantedToIdentitiesV2": [],
        "grantedToIdentities": [],
        "link": {
          "scope": "users",
          "type": "view",
          "webUrl": "https://contoso.sharepoint.com/:b:/s/pnpcoresdktestgroup/EY50lub3559MtRKfj2hrZqoBsS_o5pIcCyNIL3D_vEyG5Q",
          "preventsDownload": true
        }
      },
      {
        "id": "8bf1ca81-a63f-4796-9af5-d86ded8ce5a7",
        "roles": [
          "write"
        ],
        "hasPassword": true,
        "grantedToIdentitiesV2": [],
        "grantedToIdentities": [],
        "link": {
          "scope": "organization",
          "type": "edit",
          "webUrl": "https://contoso.sharepoint.com/:b:/s/pnpcoresdktestgroup/EY50lub3559MtRKfj2hrZqoBDyAMq6f9C2eqWwFsbei6nA",
          "preventsDownload": false
        }
      }
    ]
  };

  const graphResponseText: any = [
    {
      "id": "2a021f54-90a2-4016-b3b3-5f34d2e7d932",
      "roles": "read",
      "link": "https://contoso.sharepoint.com/:b:/s/pnpcoresdktestgroup/EY50lub3559MtRKfj2hrZqoBWnHOpGIcgi4gzw9XiWYJ-A",
      "scope": "anonymous"
    }
  ];

  const stubOdataResponse: any = (graphResponse: any = null) => {
    return sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/sites/${fileDetailsResponse.SiteId}/drives/${fileDetailsResponse.VroomDriveID}/items/${fileDetailsResponse.VroomItemID}/permissions?$filter=Link ne null`) {
        return graphResponse.value;
      }
      throw 'Invalid request';
    });
  };

  const stubOdataScopeResponse: any = (scope: any = null, graphResponse: any = null) => {
    return sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/sites/${fileDetailsResponse.SiteId}/drives/${fileDetailsResponse.VroomDriveID}/items/${fileDetailsResponse.VroomItemID}/permissions?$filter=Link ne null and Link/Scope eq '${scope}'`) {
        return graphResponse.value.filter((x: any) => x.link.scope === scope);
      }
      throw 'Invalid request';
    });
  };

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
      request.get,
      odata.getAllItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_SHARINGLINK_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves sharing links from file specified by id', async () => {
    stubOdataResponse(graphResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId, output: 'json', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse.value));
  });

  it('retrieves sharing links from file specified by url', async () => {
    stubOdataResponse(graphResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, output: 'json', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse.value));
  });

  it('retrieves sharing links from file specified by url and scope anonymous', async () => {
    const scope = 'anonymous';
    stubOdataScopeResponse(scope, graphResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, scope: scope, output: 'json', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse.value.filter(x => x.link.scope === scope)));
  });

  it('retrieves sharing links from file specified by url and scope users', async () => {
    const scope = 'users';
    stubOdataScopeResponse(scope, graphResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, scope: scope, output: 'json', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse.value.filter(x => x.link.scope === scope)));
  });

  it('retrieves sharing links from file specified by url and scope organization', async () => {
    const scope = 'organization';
    stubOdataScopeResponse(scope, graphResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, scope: scope, output: 'json', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse.value.filter(x => x.link.scope === scope)));
  });

  it('retrieves sharing links from file specified by url with output text', async () => {
    const scope = 'anonymous';
    stubOdataScopeResponse(scope, graphResponse);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, scope: scope, output: 'text', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponseText));
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
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid scope specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, scope: 'invalid scope' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
