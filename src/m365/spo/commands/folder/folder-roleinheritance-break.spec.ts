import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { formatting } from '../../../../utils/formatting';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { urlUtil } from '../../../../utils/urlUtil';
import commands from '../../commands';
import { pid } from '../../../../utils/pid';
const command: Command = require('./folder-roleinheritance-break');

describe(commands.FOLDER_ROLEINHERITANCE_BREAK, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const folderUrl = '/Shared Documents/TestFolder';
  const rootFolderUrl = '/Shared Documents';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.prompt,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      pid.getProcessName,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FOLDER_ROLEINHERITANCE_BREAK), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: folderUrl, confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and folderUrl are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderUrl: folderUrl, confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before breaking role inheritance for the folder when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderUrl: folderUrl
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts breaking role inheritance for the folder when confirm option is not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderUrl: folderUrl
      }
    });

    assert(postSpy.notCalled);
  });

  it('breaks role inheritance on folder by site-relative URL (debug)', async () => {
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, folderUrl);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/breakroleinheritance(true)`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        folderUrl: folderUrl,
        confirm: true
      }
    });
  });

  it('breaks role inheritance on folder by site-relative URL when prompt confirmed', async () => {
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, folderUrl);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/breakroleinheritance(true)`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderUrl: folderUrl
      }
    });
  });

  it('breaks role inheritance on root folder URL of a library when prompt confirmed', async () => {
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, rootFolderUrl);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')/breakroleinheritance(true)`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderUrl: rootFolderUrl
      }
    });
  });
  it('breaks role inheritance and clears existing scopes on folder by site-relative URL when prompt confirmed', async () => {
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, folderUrl);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/breakroleinheritance(false)`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderUrl: folderUrl,
        clearExistingPermissions: true
      }
    });
  });

  it('correctly handles error when breaking folder role inheritance', async () => {
    const errorMessage = 'request rejected';
    sinon.stub(request, 'post').callsFake(async () => { throw errorMessage; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        folderUrl: folderUrl,
        confirm: true
      }
    }), new CommandError(errorMessage));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});