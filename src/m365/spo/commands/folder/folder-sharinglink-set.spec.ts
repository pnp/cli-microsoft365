import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import commands from '../../commands.js';
import command from './folder-sharinglink-set.js';

describe(commands.FOLDER_SHARINGLINK_SET, () => {
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
  const id = 'ef1cddaa-b74a-4aae-8a7a-5c16b4da67f2';

  const defaultGetStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${webUrl}/_api/web/GetFolderById('${folderId}')?$select=ServerRelativeUrl`) {
        return { ServerRelativeUrl: folderUrl };
      }
      else if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter('/sites/project-x/shared documents')}')?$select=ServerRelativeUrl`) {
        return { ServerRelativeUrl: '/sites/project-x/shared documents' };
      }
      else if (opts.url === `${webUrl}/_api/web/GetFolderById('invalid')?$select=ServerRelativeUrl`) {
        throw { error: { 'odata.error': { message: { value: 'Folder Not Found.' } } } };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/project-x?$select=id`) {
        return { id: siteId };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return getDriveResponse;
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/folder1?$select=id` ||
        opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/root?$select=id`
      ) {
        return { id: itemId };
      }

      throw 'Invalid request';
    });
  };

  const getDriveResponse: any = {
    value: [
      {
        "id": driveId,
        "webUrl": `${webUrl}/Shared%20Documents`
      }
    ]
  };

  const graphResponse = {
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
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_SHARINGLINK_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderId: folderId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the folderId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: 'invalid', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the expirationDateTime option is not a valid date', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, expirationDateTime: 'invalid date', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid role specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, role: 'invalid role', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, role: 'read', id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('updates a sharing link to a folder specified by the id', async () => {
    defaultGetStub();

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions/${id}`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, folderId: folderId, role: 'read', id: id, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse));
  });

  it('updates a sharing link to a folder specified by the URL', async () => {
    defaultGetStub();

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions/${id}`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, folderUrl: '/sites/project-x/shared documents/', expirationDateTime: '2024-05-05T16:57:00.000Z', id: id, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse));
  });

  it('throws error when folder not found by id', async () => {
    defaultGetStub();

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, folderId: 'invalid', role: "read" } } as any),
      new CommandError(`Folder Not Found.`));
  });

  it('throws error when drive not found by url', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderUrl)}')?$select=ServerRelativeUrl`) {
        return { ServerRelativeUrl: folderUrl };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/project-x?$select=id`) {
        return { id: siteId };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, type: 'read' } } as any),
      new CommandError(`Drive 'https://contoso.sharepoint.com/sites/project-x/shared%20documents/folder1' not found`));
  });
});