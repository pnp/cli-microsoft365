import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './folder-sharinglink-get.js';
import { spo } from '../../../../utils/spo.js';
import { drive } from '../../../../utils/drive.js';
import { CommandError } from '../../../../Command.js';

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

  const getDriveResponse: any =
  {
    "id": driveId,
    "webUrl": `${webUrl}/Shared%20Documents`
  };

  const defaultGetStub = (): sinon.SinonStub => {
    sinon.stub(spo, 'getFolderServerRelativeUrl').resolves(folderUrl);
    sinon.stub(drive, 'getDrive').resolves(getDriveResponse);
    sinon.stub(drive, 'getDriveItemId').resolves(itemId);

    return sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions/${id}`) {
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
    sinon.stub(spo, 'getSiteId').resolves(siteId);
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
      spo.getFolderServerRelativeUrl,
      drive.getDrive,
      drive.getDriveItemId
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

  it('retrieves sharing link of folder specified by id (debug)', async () => {
    defaultGetStub();

    await command.action(logger, { options: { debug: true, webUrl: webUrl, folderId: folderId, id: id } } as any);
    assert(loggerLogSpy.calledWith(graphResponse));
  });

  it('retrieves sharing link of folder specified by url (debug)', async () => {
    defaultGetStub();

    await command.action(logger, { options: { debug: true, webUrl: webUrl, folderUrl: folderUrl, id: id } } as any);
    assert(loggerLogSpy.calledWith(graphResponse));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        folderUrl: folderUrl,
        id: id
      }
    }), new CommandError(errorMessage));
  });
});