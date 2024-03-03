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
import command from './site-hubsite-connect.js';

describe(commands.SITE_HUBSITE_CONNECT, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
    assert.strictEqual(command.name, commands.SITE_HUBSITE_CONNECT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('connects site to the hub site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site/JoinHubSite('255a50b2-527f-4413-8485-57f4c17a24d1')`) > -1) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles error when the specified id doesn\'t point to a valid hub site', async () => {
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

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } } as any),
      new CommandError('Exception of type \'Microsoft.SharePoint.Client.ResourceNotFoundException\' was thrown.'));
  });

  it('supports specifying site collection URL', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying hub site ID', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site collection URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'site.com', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the hub site ID is not a valid GUID', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
