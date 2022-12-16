import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./hubsite-register');

describe(commands.HUBSITE_REGISTER, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_REGISTER), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('registers site as a hub site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/site/RegisterHubSite`) > -1) {
        return Promise.resolve({
          "Description": null,
          "ID": "255a50b2-527f-4413-8485-57f4c17a24d1",
          "LogoUrl": "http://contoso.com/logo.png",
          "SiteId": "255a50b2-527f-4413-8485-57f4c17a24d1",
          "SiteUrl": "https://contoso.sharepoint.com/sites/sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Test site"
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "ID": "255a50b2-527f-4413-8485-57f4c17a24d1",
      "LogoUrl": "http://contoso.com/logo.png",
      "SiteId": "255a50b2-527f-4413-8485-57f4c17a24d1",
      "SiteUrl": "https://contoso.sharepoint.com/sites/sales",
      "Targets": null,
      "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
      "Title": "Test site"
    }));
  });

  it('registers site as a hub site (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/site/RegisterHubSite`) > -1) {
        return Promise.resolve({
          "Description": null,
          "ID": "255a50b2-527f-4413-8485-57f4c17a24d1",
          "LogoUrl": "http://contoso.com/logo.png",
          "SiteId": "255a50b2-527f-4413-8485-57f4c17a24d1",
          "SiteUrl": "https://contoso.sharepoint.com/sites/sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Test site"
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/sales' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "ID": "255a50b2-527f-4413-8485-57f4c17a24d1",
      "LogoUrl": "http://contoso.com/logo.png",
      "SiteId": "255a50b2-527f-4413-8485-57f4c17a24d1",
      "SiteUrl": "https://contoso.sharepoint.com/sites/sales",
      "Targets": null,
      "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
      "Title": "Test site"
    }));
  });

  it('correctly handles error when trying to register site which already is a hub site as a hub site', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-1, System.InvalidOperationException",
            "message": {
              "lang": "en-US",
              "value": "This site is already a HubSite."
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales' } } as any),
      new CommandError('This site is already a HubSite.'));
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

  it('fails validation if the specified site collection URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'site.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
