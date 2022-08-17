import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./tenant-appcatalogurl-get');

describe(commands.TENANT_APPCATALOGURL_GET, () => {
  let log: any[];
  let requests: any[];
  let logger: Logger;

  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    requests = [];
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
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
    const options = command.options;
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

    command.action(logger, {
      options: {

      }
    } as any, (err?: any) => {
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

    command.action(logger, {
      options: {
        debug: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.lastCall.args[0] === 'https://contoso.sharepoint.com/sites/apps');
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

    command.action(logger, {
      options: {
        debug: false
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);
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

    command.action(logger, {
      options: {
        debug: true
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('Tenant app catalog is not configured.'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});