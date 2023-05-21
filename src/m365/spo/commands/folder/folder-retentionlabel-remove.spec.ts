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
import * as SpoListItemRetentionLabelRemoveCommand from '../listitem/listitem-retentionlabel-remove';
import * as SpoListRetentionLabelRemoveCommand from '../list/list-retentionlabel-remove';
const command: Command = require('./folder-retentionlabel-remove');

describe(commands.FOLDER_RETENTIONLABEL_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const folderUrl = `/Shared Documents/Fo'lde'r`;
  const folderId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const listId = 1;
  const SpoListItemRetentionLabelRemoveCommandOutput = `{ "stdout": "", "stderr": "" }`;
  const SpoListRetentionLabelRemoveCommandOutput = `{ "stdout": "", "stderr": "" }`;
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
  let promptOptions: any;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => { return defaultValue; }));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.prompt,
      Cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FOLDER_RETENTIONLABEL_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing retentionlabel from a folder when confirmation argument not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, folderUrl: folderUrl } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing folder retention label when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        folderUrl: folderUrl,
        webUrl: webUrl
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the retentionlabel from a folder based on folderUrl when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderUrl)}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return folderResponse;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelRemoveCommand) {
        return ({
          stdout: SpoListItemRetentionLabelRemoveCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        folderUrl: folderUrl,
        webUrl: webUrl
      }
    }));
  });

  it('removes the retentionlabel from a folder based on folderId when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderById('${folderId}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return folderResponse;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelRemoveCommand) {
        return ({
          stdout: SpoListItemRetentionLabelRemoveCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        folderId: folderId,
        webUrl: webUrl,
        listItemId: 1
      }
    }));
  });

  it('removes the retentionlabel from a folder based on folderId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderById('${folderId}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return folderResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelRemoveCommand) {
        return ({
          stdout: SpoListItemRetentionLabelRemoveCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        folderId: folderId,
        webUrl: webUrl,
        listItemId: 1,
        confirm: true
      }
    }));
  });

  it('removes the retentionlabel to a folder if the folder is the rootfolder of a document library based on folderId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderById('${folderId}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return { ServerRelativeUrl: '/Shared Documents' };
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListRetentionLabelRemoveCommand) {
        return ({
          stdout: SpoListRetentionLabelRemoveCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        folderId: folderId,
        webUrl: webUrl,
        confirm: true
      }
    }));
  });


  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(request, 'get').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        confirm: true,
        folderUrl: folderUrl,
        webUrl: webUrl
      }
    }), new CommandError(errorMessage));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if both folderUrl or folderId options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: folderUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderUrl: folderUrl } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the folderId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the folderId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both folderId and folderUrl options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderId: folderId, folderUrl: folderUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});