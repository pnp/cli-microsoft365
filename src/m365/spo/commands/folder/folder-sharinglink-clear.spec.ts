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
import command from './folder-sharinglink-clear.js';
import { spo } from '../../../../utils/spo.js';
import { drive } from '../../../../utils/drive.js';
import { Drive } from '@microsoft/microsoft-graph-types';

describe(commands.FOLDER_SHARINGLINK_CLEAR, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

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
      }
    ]
  };

  const driveDetails: Drive = {
    id: driveId,
    webUrl: `${webUrl}/Shared%20Documents`
  };

  const getStubs: any = (options: any) => {
    sinon.stub(spo, 'getFolderServerRelativeUrl').resolves(options.folderUrl);
    sinon.stub(spo, 'getSiteIdByMSGraph').resolves(options.siteId);
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

    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation,
      odata.getAllItems,
      spo.getSiteIdByMSGraph,
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
    assert.strictEqual(command.name, commands.FOLDER_SHARINGLINK_CLEAR);
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
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId, scope: 'organization' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before clearing the sharing links from a folder when force option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderId: folderId
      }
    });

    assert(promptIssued);
  });

  it('aborts clearing the sharing links from a folder when force option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderUrl: folderUrl
      }
    });

    assert(deleteSpy.notCalled);
  });

  it('clears sharing links from folder by id for the specified scope', async () => {
    const scope = 'anonymous';
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });
    stubOdataScopeResponse(scope, graphResponse);

    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions/2a021f54-90a2-4016-b3b3-5f34d2e7d932`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        folderId: folderId,
        scope: scope
      }
    });

    assert(requestDeleteStub.called);
  });

  it('clears sharing links from folder by URL for the all scopes', async () => {
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });
    stubOdataResponse(graphResponse);

    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions/2a021f54-90a2-4016-b3b3-5f34d2e7d932`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        folderUrl: folderUrl,
        force: true
      }
    });

    assert(requestDeleteStub.called);
  });

  it('throws error when drive not found by url', async () => {
    sinon.stub(spo, 'getFolderServerRelativeUrl').resolves(folderUrl);
    sinon.stub(spo, 'getSiteIdByMSGraph').resolves(siteId);
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, force: true } } as any),
      new CommandError(`Drive 'https://contoso.sharepoint.com/sites/project-x/shared%20documents/folder1' not found`));
  });
});