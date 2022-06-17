import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./hubsite-register');

describe(commands.HUBSITE_REGISTER, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_REGISTER), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('registers site as a hub site', (done) => {
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

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('registers site as a hub site (debug)', (done) => {
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

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/sales' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to register site which already is a hub site as a hub site', (done) => {
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

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('This site is already a HubSite.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying site collection URL', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site collection URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'site.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});