import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './folder-sharinglink-get.js';

describe(commands.FOLDER_SHARINGLINK_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const folderUrl = '/sites/project-x/shared documents/folder1';
  const folderId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';
  const id = 'd6f6a428-9857-471f-9635-edd68d5aa6c1';
  const siteId = '0f9b8f4f-0e8e-4630-bb0a-501442db9b64';
  const driveId = '013TMHP6UOOSLON57HT5GLKEU7R5UGWZVK';
  const itemId = 'b!T4-bD44OMEa7ClAUQtubZID9tc40pGJKpguycvELod_Gx-lo4ZQiRJ7vylonTufG';

  const graphResponse = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#drives('b!T4-bD44OMEa7ClAUQtubZID9tc40pGJKpguycvELod_Gx-lo4ZQiRJ7vylonTufG')/items('013TMHP6UOOSLON57HT5GLKEU7R5UGWZVK')/permissions/$entity",
    "@deprecated.GrantedToIdentities": "GrantedToIdentities has been deprecated. Refer to GrantedToIdentitiesV2",
    "id": "d6f6a428-9857-471f-9635-edd68d5aa6c1",
    "roles": [
      "write"
    ],
    "shareId": "u!aHR0cHM6Ly9uYWNoYW4zNjUuc2hhcmVwb2ludC5jb20vOmY6L3MvU1BEZW1vL0VxXzlYcXRIdks1RW9wd3NfX1kteko0QlNybFFNUy1qUTBFOWJsazhVLVNTdVE",
    "hasPassword": false,
    "grantedToIdentitiesV2": [],
    "grantedToIdentities": [],
    "link": {
      "scope": "anonymous",
      "type": "edit",
      "webUrl": "https://contoso.sharepoint.com/:f:/s/project-x/Eq_9XqtHvK5Eopws__Y-zJ4BSrlQMS-jQ0E9blk8U-SSuQ",
      "preventsDownload": false
    }
  };

  const getDriveResponse: any = {
    value: [
      {
        "id": driveId,
        "webUrl": `${webUrl}/Shared%20Documents`
      }
    ]
  };

  const defaultGetStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/project-x?$select=id`) {
        return { id: siteId };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`) {
        return getDriveResponse;
      }
      else if (opts.url === `${webUrl}/_api/web/GetFolderById('${folderId}')?$select=ServerRelativeUrl`) {
        return { ServerRelativeUrl: folderUrl };
      }
      else if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter('/sites/project-x/shared documents/folder1')}')?$select=ServerRelativeUrl`) {
        return { ServerRelativeUrl: '/sites/project-x/shared documents' };
      }
      else if (opts.url === `${webUrl}/_api/web/GetFolderById('invalid')?$select=ServerRelativeUrl`) {
        throw { error: { 'odata.error': { message: { value: 'File Not Found.' } } } };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/folder1?$select=id` ||
        opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/root?$select=id`) {
        return { id: itemId };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions/${id}`) {
        return graphResponse;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_SHARINGLINK_GET);
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

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId, id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId, id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves sharing link of folder specified by id', async () => {
    defaultGetStub();

    await command.action(logger, { options: { webUrl: webUrl, folderId: folderId, id: id } } as any);
    assert(loggerLogSpy.calledWith(graphResponse));
  });

  it('retrieves sharing link of folder specified by url', async () => {
    defaultGetStub();

    await command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, id: id } } as any);
    assert(loggerLogSpy.calledWith(graphResponse));
  });

  it('throws error when folder not found by id', async () => {
    defaultGetStub();

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, folderId: 'invalid', id: id, verbose: true } } as any),
      new CommandError(`File Not Found.`));
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

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, id: id, verbose: true } } as any),
      new CommandError(`Drive 'https://contoso.sharepoint.com/sites/project-x/shared%20documents/folder1' not found`));
  });
});