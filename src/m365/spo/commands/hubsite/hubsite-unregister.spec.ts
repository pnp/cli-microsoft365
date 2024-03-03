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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './hubsite-unregister.js';

describe(commands.HUBSITE_UNREGISTER, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
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
    requests = [];
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
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
    assert.strictEqual(command.name, commands.HUBSITE_UNREGISTER);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('unregisters the specified hub site without prompting with confirmation argument', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/site/UnregisterHubSite' &&
        opts.headers &&
        opts.headers.accept &&
        (opts.headers.accept as string).indexOf('application/json') === 0) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/sales', force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('prompts before unregistering the hub site when confirmation argument not passed', async () => {
    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/sales' } });

    assert(promptIssued);
  });

  it('aborts unregistering hub site when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/sales' } });
    assert(requests.length === 0);
  });

  it('unregisters hub site when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/site/UnregisterHubSite' &&
        opts.headers &&
        opts.headers.accept &&
        (opts.headers.accept as string).indexOf('application/json') === 0) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/sales' } });
  });

  it('correctly handles failure when the specified site is not a hub site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/site/UnregisterHubSite') {
        throw {
          error: {
            "odata.error": {
              "code": "-2147024809, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "hubSiteId"
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/sales', force: true } } as any),
      new CommandError("hubSiteId"));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
