import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { spo } from '../../../../utils/spo.js';
import { drive } from '../../../../utils/drive.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import commands from '../../commands.js';
import command from './folder-sharinglink-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.FOLDER_SHARINGLINK_ADD, () => {
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

  const driveDetails: Drive = {
    id: driveId,
    webUrl: `${webUrl}/Shared%20Documents`
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

  const getStubs: any = (options: any) => {
    sinon.stub(spo, 'getFolderServerRelativeUrl').resolves(options.folderUrl);
    sinon.stub(spo, 'getSiteId').resolves(options.siteId);
    sinon.stub(drive, 'getDriveByUrl').resolves(options.drive);
    sinon.stub(drive, 'getDriveItemId').resolves(options.itemId);
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => settingName === settingsNames.prompt ? false : defaultValue);
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
      request.post,
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
    assert.strictEqual(command.name, commands.FOLDER_SHARINGLINK_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderId: folderId, type: "view" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the folderId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: 'invalid', type: "view" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the expirationDateTime option is not a valid date', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, expirationDateTime: 'invalid date', type: 'view' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid scope specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, scope: 'invalid scope', type: 'view' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid type specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, type: 'invalid type' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is users but recipients are not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, type: 'view', scope: 'users' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if recipients option is not a valid UPN', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, type: 'view', scope: 'users', recipients: "invalid upn" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderId: folderId, type: 'view' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('creates a sharing link to a folder specified by the id', async () => {
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/createLink`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, folderId: folderId, type: 'view', scope: 'organization', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse));
  });

  it('creates a sharing link to a folder specified by the URL', async () => {
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/createLink`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, type: 'view', scope: 'users', recipients: 'john@contoso.com', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(graphResponse));
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

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, type: 'view' } } as any),
      new CommandError(`Drive 'https://contoso.sharepoint.com/sites/project-x/shared%20documents/folder1' not found`));
  });
});