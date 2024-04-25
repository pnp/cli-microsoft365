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
import commands from '../../commands.js';
import command from './folder-sharinglink-remove.js';
import { spo } from '../../../../utils/spo.js';
import { drive } from '../../../../utils/drive.js';
import { Drive } from '@microsoft/microsoft-graph-types';

describe(commands.FOLDER_SHARINGLINK_REMOVE, () => {
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
  const id = 'ef1cddaa-b74a-4aae-8a7a-5c16b4da67f2';

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

    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation,
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
    assert.strictEqual(command.name, commands.FOLDER_SHARINGLINK_REMOVE);
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

  it('passes validation if options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId, id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified sharing link to a folder when force option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderId: folderId,
        id: id
      }
    });

    assert(promptIssued);
  });

  it('aborts removing the specified sharing link to a folder when force option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderUrl: folderUrl,
        id: id
      }
    });

    assert(deleteSpy.notCalled);
  });

  it('removes specified sharing link to a folder by folderId when prompt confirmed', async () => {
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });

    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions/${id}`) {
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
        id: id
      }
    });
    assert(requestDeleteStub.called);
  });

  it('removes specified sharing link to a folder by folderUrl when prompt confirmed', async () => {
    getStubs({ folderUrl: folderUrl, siteId: siteId, drive: driveDetails, itemId: itemId });

    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions/${id}`) {
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
        folderUrl: folderUrl,
        id: id,
        force: true
      }
    });
    assert(requestDeleteStub.called);
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

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl, force: true } } as any),
      new CommandError(`Drive 'https://contoso.sharepoint.com/sites/project-x/shared%20documents/folder1' not found`));
  });
});