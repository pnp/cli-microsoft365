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
import command, { options } from './site-hubsite-theme-sync.js';

describe(commands.SITE_HUBSITE_THEME_SYNC, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
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
    assert.strictEqual(command.name, commands.SITE_HUBSITE_THEME_SYNC);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('syncs hub site theme to a web', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/SyncHubSiteTheme`) > -1) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/Project-X' }) });
    assert(loggerLogSpy.notCalled);
  });

  it('syncs hub site theme to a web (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/SyncHubSiteTheme`) > -1) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' }) });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles error when hub site not found', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        "odata.error": {
          "code": "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
          "message": {
            "lang": "en-US",
            "value": "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/Project-X' }) }),
      new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'Invalid' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'Invalid' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passed validation if webUrl is a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });
});
