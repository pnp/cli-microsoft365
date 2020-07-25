import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./hubsite-connect');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.HUBSITE_CONNECT, () => {
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
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_CONNECT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('connects site to the hub site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/site/JoinHubSite('255a50b2-527f-4413-8485-57f4c17a24d1')`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', hubSiteId: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('connects site to the hub site (true)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/site/JoinHubSite('255a50b2-527f-4413-8485-57f4c17a24d1')`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/sales', hubSiteId: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified id doesn\'t point to a valid hub site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
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
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', hubSiteId: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Exception of type \'Microsoft.SharePoint.Client.ResourceNotFoundException\' was thrown.')));
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

  it('supports specifying hub site ID', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--hubSiteId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site collection URL is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'site.com', hubSiteId: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the hub site ID is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', hubSiteId: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', hubSiteId: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert.strictEqual(actual, true);
  });
});