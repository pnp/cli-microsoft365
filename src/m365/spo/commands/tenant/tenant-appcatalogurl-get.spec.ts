import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./tenant-appcatalogurl-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TENANT_APPCATALOGURL_GET, () => {
  let log: any[];
  let requests: any[];
  let cmdInstance: any;

  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    requests = [];
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
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_APPCATALOGURL_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('handles promise error while getting tenant appcatalog', (done) => {
    // get tenant app catalog
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {

      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the tenant appcatalog url (debug)', (done) => {
    // get tenant app catalog
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve(JSON.stringify({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" }));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.args[0] === 'https://contoso.sharepoint.com/sites/apps');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles if tenant appcatalog is null or not exist', (done) => {
    // get tenant app catalog
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve(JSON.stringify({ "CorporateCatalogUrl": null }));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles if tenant appcatalog is null or not exist (debug)', (done) => {
    // get tenant app catalog
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve(JSON.stringify({ "CorporateCatalogUrl": null }));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Tenant app catalog is not configured.'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});