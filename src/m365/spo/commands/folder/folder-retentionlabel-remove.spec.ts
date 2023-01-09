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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as SpoListItemRetentionLabelRemoveCommand from '../listitem/listitem-retentionlabel-remove';
const command: Command = require('./folder-retentionlabel-remove');

describe(commands.FOLDER_RETENTIONLABEL_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const folderUrl = `/Shared Documents/Fo'lde'r`;
  const folderId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const listId = 1;
  const SpoListItemRetentionLabelRemoveCommandOutput = `{ "stdout": "", "stderr": "" }`;

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.prompt,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FOLDER_RETENTIONLABEL_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing retentionlabel from a folder when confirmation argument not passed', async () => {
    await command.action(logger, { options: { debug: false, webUrl: webUrl, folderUrl: folderUrl } });
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
        debug: false,
        folderUrl: folderUrl,
        webUrl: webUrl
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the retentionlabel from a folder based on folderUrl when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderUrl)}')?$expand=ListItemAllFields`) {
        return { ListItemAllFields: { Id: listId }, ServerRelativeUrl: folderUrl };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    const postSpy = sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelRemoveCommand) {
        return ({
          stdout: SpoListItemRetentionLabelRemoveCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: false,
        folderUrl: folderUrl,
        webUrl: webUrl
      }
    });
    assert(postSpy.called);
  });

  it('removes the retentionlabel from a folder based on folderId when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderById('${folderId}')?$expand=ListItemAllFields`) {
        return { ListItemAllFields: { Id: listId }, ServerRelativeUrl: folderUrl };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    const postSpy = sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelRemoveCommand) {
        return ({
          stdout: SpoListItemRetentionLabelRemoveCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: false,
        folderId: folderId,
        webUrl: webUrl,
        listItemId: 1
      }
    });
    assert(postSpy.called);
  });

  it('removes the retentionlabel from a folder based on folderId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderById('${folderId}')?$expand=ListItemAllFields`) {
        return { ListItemAllFields: { Id: listId }, ServerRelativeUrl: folderUrl };
      }

      throw 'Invalid request';
    });

    const postSpy = sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelRemoveCommand) {
        return ({
          stdout: SpoListItemRetentionLabelRemoveCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: false,
        folderId: folderId,
        webUrl: webUrl,
        listItemId: 1,
        confirm: true
      }
    });
    assert(postSpy.called);
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