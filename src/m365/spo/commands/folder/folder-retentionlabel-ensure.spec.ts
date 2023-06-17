import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as SpoListItemRetentionLabelEnsureCommand from '../listitem/listitem-retentionlabel-ensure';
import * as SpoListRetentionLabelEnsureCommand from '../list/list-retentionlabel-ensure';
const command: Command = require('./folder-retentionlabel-ensure');

describe(commands.FOLDER_RETENTIONLABEL_ENSURE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const folderUrl = `/Shared Documents/Fo'lde'r`;
  const folderId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const listId = 1;
  const retentionlabelName = "retentionlabel";
  const SpoListItemRetentionLabelEnsureCommandOutput = `{ "stdout": "", "stderr": "" }`;
  const SpoListRetentionLabelEnsureCommandOutput = `{ "stdout": "", "stderr": "" }`;
  const folderResponse = {
    ListItemAllFields: {
      Id: listId,
      ParentList: {
        Id: '75c4d697-bbff-40b8-a740-bf9b9294e5aa'
      }
    }
  };

  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_RETENTIONLABEL_ENSURE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds the retentionlabel to a folder based on folderUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderUrl)}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return folderResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelEnsureCommand) {
        return ({
          stdout: SpoListItemRetentionLabelEnsureCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        folderUrl: folderUrl,
        webUrl: webUrl,
        name: retentionlabelName
      }
    }));
  });

  it('adds the retentionlabel to a folder based on folderId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderById('${folderId}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return folderResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelEnsureCommand) {
        return ({
          stdout: SpoListItemRetentionLabelEnsureCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        folderId: folderId,
        webUrl: webUrl,
        name: retentionlabelName
      }
    }));
  });

  it('adds the retentionlabel to a folder if the folder is the rootfolder of a document library based on folderId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderById('${folderId}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return { ServerRelativeUrl: '/Shared Documents' };
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListRetentionLabelEnsureCommand) {
        return ({
          stdout: SpoListRetentionLabelEnsureCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        folderId: folderId,
        webUrl: webUrl,
        name: retentionlabelName
      }
    }));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: retentionlabelName,
        folderUrl: folderUrl,
        webUrl: webUrl
      }
    }), new CommandError(errorMessage));
  });

  it('fails validation if both folderUrl or folderId options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: folderUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderUrl: folderUrl, name: retentionlabelName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the folderId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: '12345', name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the folderId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId, name: retentionlabelName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both folderId and folderUrl options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId, folderUrl: folderUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});