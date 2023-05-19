import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./folder-roleinheritance-reset');

describe(commands.FOLDER_ROLEINHERITANCE_RESET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const folderUrl = 'Shared Documents/TestFolder';
  const rootFolderUrl = '/Shared Documents';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_ROLEINHERITANCE_RESET);
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

  it('prompts before resetting role inheritance for the folder when confirm option not passed', async () => {
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

  it('aborts resetting role inheritance for the folder when confirm option is not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        folderUrl: folderUrl
      }
    });

    assert(postSpy.notCalled);
  });

  it('resets role inheritance on folder by site-relative URL (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%20Documents%2FTestFolder')/ListItemAllFields/resetroleinheritance`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        folderUrl: folderUrl,
        confirm: true
      }
    });
  });

  it('resets role inheritance on folder by site-relative URL when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%20Documents%2FTestFolder')/ListItemAllFields/resetroleinheritance`) {
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

  it('resets role inheritance on root folder URL of a library when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetList('%2Fsites%2Fproject-x%2FShared%20Documents')/resetroleinheritance`) {
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

  it('correctly handles error when resetting folder role inheritance', async () => {
    const errorMessage = 'Cannot find resource';
    sinon.stub(request, 'post').callsFake(async () => {
      throw {
        error: {
          'odata.error': {
            message: {
              value: errorMessage
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        folderUrl: folderUrl,
        confirm: true
      }
    }), new CommandError(errorMessage));
  });
});