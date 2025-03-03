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
import command from './page-remove.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.PAGE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;
  let postStub: sinon.SinonStub;
  let deleteStub: sinon.SinonStub;
  let promptForConfirmationStub: sinon.SinonStub;

  const webUrl = 'https://contoso.sharepoint.com/sites/Marketing';
  const serverRelativeUrl = urlUtil.getServerRelativeSiteUrl(webUrl);
  const pageName = 'HR.aspx';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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

    promptForConfirmationStub = sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    const serverRelativePageUrl = `${serverRelativeUrl}/SitePages/${pageName}`;
    postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePageUrl)}')/Recycle`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePageUrl)}')`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should prompt before removing page when confirmation argument not passed', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName
        }
      });

    assert(promptIssued);
  });

  it('aborts removing page when prompt not confirmed', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName
        }
      });

    assert(deleteStub.notCalled);
    assert(postStub.notCalled);
  });

  it('logs no command output', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName,
          force: true
        }
      });

    assert(loggerLogSpy.notCalled);
  });

  it('permanently removes page when prompt confirmed', async () => {
    promptForConfirmationStub.restore();
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName,
          verbose: true
        }
      });

    assert(deleteStub.calledOnce);
    assert(postStub.notCalled);
  });

  it('correctly recycles page', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName,
          recycle: true,
          force: true
        }
      });

    assert(postStub.calledOnce);
    assert(deleteStub.notCalled);
  });

  it('correctly bypasses a shared lock', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName,
          bypassSharedLock: true,
          force: true
        }
      });

    assert.deepStrictEqual(deleteStub.firstCall.args[0].headers.Prefer, 'bypass-shared-lock');
  });

  it('correctly recycles a page when extension is not specified', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName.substring(0, pageName.lastIndexOf('.')),
          recycle: true,
          force: true
        }
      });

    assert(postStub.calledOnce);
    assert(deleteStub.notCalled);
  });

  it('correctly removes a nested page', async () => {
    const pageUrl = '/folder1/folder2/' + pageName;
    deleteStub.restore();

    deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl + '/SitePages' + pageUrl)}')`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageUrl,
          force: true
        }
      });

    assert(deleteStub.calledOnce);
  });

  it('correctly handles API error', async () => {
    deleteStub.restore();
    const errorMessage = 'The file /sites/Marketing/SitePages/My-new-page.aspx does not exist.';

    sinon.stub(request, 'delete').rejects({
      error: {
        'odata.error': {
          message: {
            lang: 'en-US',
            value: errorMessage
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, name: pageName, force: true } }),
      new CommandError(errorMessage));
  });

  it('fails validation if webUrl is not an absolute URL', async () => {
    const actual = await command.validate({ options: { name: pageName, webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({
      options: { name: pageName, webUrl: 'http://foo' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({
      options: { name: pageName, webUrl: webUrl }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name has no extension', async () => {
    const actual = await command.validate({
      options: { name: 'page', webUrl: webUrl }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
