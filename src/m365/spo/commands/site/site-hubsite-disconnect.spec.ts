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
import command, { options } from './site-hubsite-disconnect.js';

describe(commands.SITE_HUBSITE_DISCONNECT, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
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
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_HUBSITE_DISCONNECT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('disconnects the site from its hub site without prompting for confirmation when force option specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_api/site/JoinHubSite('00000000-0000-0000-0000-000000000000')`) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/Sales', force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('disconnects the site from its hub site without prompting for confirmation when force option specified (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_api/site/JoinHubSite('00000000-0000-0000-0000-000000000000')`) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/Sales', force: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('disconnects site from the hub site as admin', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/ConnectSiteToHubSiteById`) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', asAdmin: true, force: true } });
    assert(postStub.calledOnce);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { siteUrl: 'https://contoso.sharepoint.com/sites/sales', hubSiteId: '00000000-0000-0000-0000-000000000000' });
  });

  it('prompts before disconnecting the specified site from its hub site when force option not passed', async () => {
    sinon.stub(request, 'post').resolves();
    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/Sales' } });

    assert(promptIssued);
  });

  it('aborts disconnecting site from its hub site when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/Sales' } });
    assert(postSpy.notCalled);
  });

  it('disconnects the site from its hub site when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async () => {
      return ({
        "odata.null": true
      });
    });
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/Sales' } });
    assert(postStub.called);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw {
        error: {
          "odata.error": {
            "code": "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
            "message": {
              "lang": "en-US",
              "value": "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', force: true } } as any),
      new CommandError('Exception of type \'Microsoft.SharePoint.Client.ResourceNotFoundException\' was thrown.'));
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'abc' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation when url is a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com/sites/Sales' });
    assert.strictEqual(actual.success, true);
  });
});
