import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./hubsite-register');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.HUBSITE_REGISTER, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
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

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
    sinon.stub(request, 'post').callsFake((opts) => {
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

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales' } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying site collection URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site collection URL is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'site.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales' } });
    assert.strictEqual(actual, true);
  });
});