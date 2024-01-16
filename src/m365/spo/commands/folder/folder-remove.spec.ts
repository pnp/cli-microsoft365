import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './folder-remove.js';

describe(commands.FOLDER_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;
  let stubPost: sinon.SinonStub;

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

    stubPost = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url!.indexOf('/_api/web/GetFolderByServerRelativePath(DecodedUrl=') >= 0) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
    requests = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing folder when confirmation argument not passed', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', url: '/Shared Documents' } });

    assert(promptIssued);
  });

  it('aborts removing folder when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', url: '/Shared Documents' } });
    assert(requests.length === 0);
  });

  it('removes the folder when prompt confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents/Folder1'
      }
    });
    assert(stubPost.called);
  });

  it('should send params for remove request', async () => {
    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents/Folder1',
        force: true
      }
    });
    const lastCall: any = stubPost.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolder1\')');
    assert.strictEqual(lastCall.method, 'POST');
    assert.strictEqual(lastCall.headers['X-HTTP-Method'], 'DELETE');
  });

  it('should send params for remove request for sites/test1', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/test1',
        url: '/Shared Documents/Folder1'
      }
    });
    const lastCall: any = stubPost.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/sites/test1/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2Fsites%2Ftest1%2FShared%20Documents%2FFolder1\')');
    assert.strictEqual(lastCall.method, 'POST');
    assert.strictEqual(lastCall.headers['X-HTTP-Method'], 'DELETE');
  });

  it('should send params for recycle request when recycle is set to true', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options:
      {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents/Folder1',
        recycle: true
      }
    });
    const lastCall: any = stubPost.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolder1\')/recycle()');
    assert.strictEqual(lastCall.method, 'POST');
    assert.strictEqual(lastCall.headers['X-HTTP-Method'], 'DELETE');
  });

  it('should show error on request reject', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };

    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').rejects(error);

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, {
      options:
      {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents/Folder1',
        recycle: true
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', url: '/Shared Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and url specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', url: '/Shared Documents' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
