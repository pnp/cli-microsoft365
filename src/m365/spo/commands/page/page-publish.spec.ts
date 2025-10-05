import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command, { options } from './page-publish.js';

describe(commands.PAGE_PUBLISH, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let postStub: sinon.SinonStub;
  let commandOptionsSchema: typeof options;

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
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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

    const serverRelativePageUrl = `${serverRelativeUrl}/SitePages/${pageName}`;
    postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePageUrl)}')/Publish()`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_PUBLISH);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('logs no command output', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName
        }
      });

    assert(loggerLogSpy.notCalled);
  });

  it('correctly publishes page', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName
        }
      });

    assert(postStub.calledOnce);
  });

  it('correctly publishes a page when extension is not specified', async () => {
    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageName.substring(0, pageName.lastIndexOf('.')),
          verbose: true
        }
      });

    assert(postStub.calledOnce);
  });

  it('correctly publishes a nested page', async () => {
    const pageUrl = '/folder1/folder2/' + pageName;
    postStub.restore();

    postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl + '/SitePages' + pageUrl)}')/Publish()`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger,
      {
        options: {
          webUrl: webUrl,
          name: pageUrl
        }
      });

    assert(postStub.calledOnce);
  });

  it('correctly handles API error', async () => {
    postStub.restore();
    const errorMessage = 'The file /sites/Marketing/SitePages/My-new-page.aspx does not exist.';

    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          message: {
            lang: 'en-US',
            value: errorMessage
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, name: pageName } }),
      new CommandError(errorMessage));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when the webUrl is a valid SharePoint URL and name is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, name: pageName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when name has no extension', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, name: 'page' });
    assert.strictEqual(actual.success, true);
  });
});
