import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
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
import command from './orgnewssite-remove.js';

describe(commands.ORGNEWSSITE_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
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
    auth.service.active = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ORGNEWSSITE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('completes a remove request - confirm parameter', async () => {
    const svcListRequest = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest']) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "8992299e-a003-4000-7686-fda36e26a53c" }, 22, []]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        force: true,
        url: "http://contoso.sharepoint.com/sites/site1"
      }
    } as any);
    assert(svcListRequest.called);
  });

  it('completes a remove request - prompt confirmed', async () => {
    const svcListRequest = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest']) {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "8992299e-a003-4000-7686-fda36e26a53c" }, 22, []]);
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        verbose: true,
        force: false,
        url: "http://contoso.sharepoint.com/sites/site1"
      }
    } as any);
    assert(svcListRequest.called);
  });

  it('handles error during remove request', async () => {
    const svcListRequest = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest']) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
              }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        force: true
      }
    } as any), new CommandError('An error has occurred'));
    assert(svcListRequest.called);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/site1', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('prompts before removing', async () => {
    await command.action(logger, { options: { debug: true, verbose: true, force: false, url: 'https://contoso.sharepoint.com/sites/test1' } });

    assert(promptIssued);
  });

  it('aborts when declined confirmation', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(() => {
      throw 'Invalid request';
    });
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { debug: true, verbose: true, force: false, url: 'https://contoso.sharepoint.com/sites/test1' } });
    assert(postStub.notCalled);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert(actual);
  });

  it('supports suppressing confirmation prompt', () => {
    const options = command.options;
    let containsConfirmOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--force') > -1) {
        containsConfirmOption = true;
      }
    });
    assert(containsConfirmOption);
  });
});
