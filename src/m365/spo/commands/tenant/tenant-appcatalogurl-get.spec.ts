import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
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
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent,
      pid.getProcessName
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

  it('handles promise error while getting tenant appcatalog', async () => {
    // get tenant app catalog
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('gets the tenant appcatalog url (debug)', async () => {
    // get tenant app catalog
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve(JSON.stringify({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" }));
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert(loggerLogSpy.lastCall.args[0] === 'https://contoso.sharepoint.com/sites/apps');
  });

  it('handles if tenant appcatalog is null or not exist', async () => {
    // get tenant app catalog
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve(JSON.stringify({ "CorporateCatalogUrl": null }));
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
      }
    });
  });

  it('handles if tenant appcatalog is null or not exist (debug)', async () => {
    // get tenant app catalog
    sinon.stub(request, 'get').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve(JSON.stringify({ "CorporateCatalogUrl": null }));
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert(loggerLogToStderrSpy.calledWith('Tenant app catalog is not configured.'));
  });
});