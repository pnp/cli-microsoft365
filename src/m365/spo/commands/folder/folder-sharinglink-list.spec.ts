import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import commands from '../../commands.js';
import command from './folder-sharinglink-list.js';
import { spo } from '../../../../utils/spo.js';
import { drive } from '../../../../utils/drive.js';
import { Drive } from '@microsoft/microsoft-graph-types';

describe(commands.FOLDER_SHARINGLINK_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const folderId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';
  const folderUrl = '/sites/project-x/shared documents/folder1';
  const siteId = '0f9b8f4f-0e8e-4630-bb0a-501442db9b64';
  const driveId = '013TMHP6UOOSLON57HT5GLKEU7R5UGWZVK';
  const itemId = 'b!T4-bD44OMEa7ClAUQtubZID9tc40pGJKpguycvELod_Gx-lo4ZQiRJ7vylonTufG';

  const graphResponse = {
    value: [
      {
        "id": "2a021f54-90a2-4016-b3b3-5f34d2e7d932",
        "roles": [
          "read"
        ],
        "shareId": "u!aHR0cHM6Ly83NTY2YXZhLnNoYXJlcG9pbnQuY29tLzpmOi9nL0V2QVFpdnpLV2ZoT3ZJOHJiNm1UVEhjQnZ4SFBWWW10aGRJNUpYdG51cGhOeUE",
        "hasPassword": false,
        "grantedToIdentitiesV2": [],
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
        "shareId": "u!aHR0cHM6Ly83NTY2YXZhLnNoYXJlcG9pbnQuY29tLzpmOi9nL0V2QVFpdnpLV2ZoT3ZJOHJiNm1UVEhjQmQzUUxCOXVsUGIyQTQ1UG81ZmRuYWc",
        "hasPassword": true,
        "grantedToIdentitiesV2": [],
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

  const driveDetails: Drive = {
    id: driveId,
    webUrl: `${webUrl}/Shared%20Documents`
  };

  const getStubs: any = (options: any) => {
    sinon.stub(spo, 'getFolderServerRelativeUrl').resolves(options.folderUrl);
    sinon.stub(spo, 'getSiteId').resolves(options.siteId);
    sinon.stub(drive, 'getDriveByUrl').resolves(options.drive);
    sinon.stub(drive, 'getDriveItemId').resolves(options.itemId);
  };

  const stubOdataResponse: any = (graphResponse: any = null) => {
    return sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions?$filter=Link ne null`) {
        return graphResponse.value;
      }
      throw 'Invalid request';
    });
  };

  const stubOdataScopeResponse: any = (scope: any = null, graphResponse: any = null) => {
    return sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions?$filter=Link ne null and Link/Scope eq '${scope}'`) {
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
      odata.getAllItems,
      spo.getSiteId,
      spo.getFolderServerRelativeUrl,
      drive.getDriveByUrl,
      drive.getDriveItemId
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_SHARINGLINK_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderId: folderId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the folderId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid scope specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId, scope: 'invalid scope' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves sharing links from folder specified by id', async () => {
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });
    stubOdataResponse(graphResponse);

    await command.action(logger, { options: { webUrl: webUrl, folderId: folderId, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse.value));
  });

  it('retrieves sharing links from folder specified by url', async () => {
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });
    stubOdataResponse(graphResponse);

    await command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse.value));
  });

  it('retrieves sharing links from folder specified by id and valid scope', async () => {
    const scope = 'organization';
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });
    stubOdataScopeResponse(scope, graphResponse);

    await command.action(logger, { options: { webUrl: webUrl, folderId: folderId, scope: scope } } as any);
    assert(loggerLogSpy.calledWith(graphResponse.value.filter(x => x.link.scope === scope)));
  });

  it('retrieves sharing links from folder specified by id and output as text', async () => {
    const scope = 'anonymous';
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });
    stubOdataScopeResponse(scope, graphResponse);

    await command.action(logger, { options: { webUrl: webUrl, folderId: folderId, scope: scope, output: 'text' } } as any);
    assert(loggerLogSpy.calledWith(graphResponseText));
  });

  it('throws error when drive not found by url', async () => {
    sinon.stub(spo, 'getFolderServerRelativeUrl').resolves(folderUrl);
    sinon.stub(spo, 'getSiteId').resolves(siteId);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, verbose: true } } as any),
      new CommandError(`Drive 'https://contoso.sharepoint.com/sites/project-x/shared%20documents/folder1' not found`));
  });
});