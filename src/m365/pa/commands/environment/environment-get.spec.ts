import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./environment-get');

describe(commands.ENVIRONMENT_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ENVIRONMENT_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'id', 'location', 'displayName', 'provisioningState', 'environmentSku', 'azureRegionHint', 'isDefault']);
  });

  it('retrieves information about the specified environment (debug)', (done) => {
    const env: any = { "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "location": "europe", "type": "Microsoft.PowerApps/environments", "id": "/providers/Microsoft.PowerApps/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "properties": { "displayName": "Contoso (default)", "createdTime": "2018-03-22T20:20:46.08653Z", "createdBy": { "id": "SYSTEM", "displayName": "SYSTEM", "type": "NotSpecified" }, "provisioningState": "Succeeded", "creationType": "DefaultTenant", "environmentSku": "Default", "environmentType": "Production", "isDefault": true, "azureRegionHint": "westeurope", "runtimeEndpoints": { "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com", "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com", "microsoft.PowerApps": "https://europe.api.powerapps.com", "microsoft.Flow": "https://europe.api.flow.microsoft.com" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(env);
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(env));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified environment', (done) => {
    const env: any = { "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "location": "europe", "type": "Microsoft.PowerApps/environments", "id": "/providers/Microsoft.PowerApps/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "properties": { "displayName": "Contoso (default)", "createdTime": "2018-03-22T20:20:46.08653Z", "createdBy": { "id": "SYSTEM", "displayName": "SYSTEM", "type": "NotSpecified" }, "provisioningState": "Succeeded", "creationType": "DefaultTenant", "environmentSku": "Default", "environmentType": "Production", "isDefault": true, "azureRegionHint": "westeurope", "runtimeEndpoints": { "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com", "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com", "microsoft.PowerApps": "https://europe.api.powerapps.com", "microsoft.Flow": "https://europe.api.flow.microsoft.com" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(env);
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(env));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
        }
      });
    });

    command.action(logger, { options: { debug: false, name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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

  it('supports specifying name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
