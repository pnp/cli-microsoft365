import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './environment-get.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.ENVIRONMENT_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
    auth.connection.active = true;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENVIRONMENT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the default environment', async () => {
    const env: any = { "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "location": "europe", "type": "Microsoft.PowerApps/environments", "id": "/providers/Microsoft.PowerApps/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "properties": { "displayName": "Contoso (default)", "createdTime": "2018-03-22T20:20:46.08653Z", "createdBy": { "id": "SYSTEM", "displayName": "SYSTEM", "type": "NotSpecified" }, "provisioningState": "Succeeded", "creationType": "DefaultTenant", "environmentSku": "Default", "environmentType": "Production", "isDefault": true, "azureRegionHint": "westeurope", "runtimeEndpoints": { "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com", "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com", "microsoft.PowerApps": "https://europe.api.powerapps.com", "microsoft.Flow": "https://europe.api.flow.microsoft.com" } } };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://api.powerapps.com/providers/Microsoft.PowerApps/environments/~default?api-version=2016-11-01') {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return env;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(env));
  });

  it('retrieves information about the specified environment', async () => {
    const env: any = { "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "location": "europe", "type": "Microsoft.PowerApps/environments", "id": "/providers/Microsoft.PowerApps/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5", "properties": { "displayName": "Contoso (default)", "createdTime": "2018-03-22T20:20:46.08653Z", "createdBy": { "id": "SYSTEM", "displayName": "SYSTEM", "type": "NotSpecified" }, "provisioningState": "Succeeded", "creationType": "DefaultTenant", "environmentSku": "Default", "environmentType": "Production", "isDefault": true, "azureRegionHint": "westeurope", "runtimeEndpoints": { "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com", "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com", "microsoft.PowerApps": "https://europe.api.powerapps.com", "microsoft.Flow": "https://europe.api.flow.microsoft.com" } } };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://api.powerapps.com/providers/Microsoft.PowerApps/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5?api-version=2016-11-01') {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return env;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert(loggerLogSpy.calledWith(env));
  });

  it('correctly handles no environment found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
      }
    });

    await assert.rejects(command.action(logger, { options: { name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } } as any),
      new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } } as any),
      new CommandError('An error has occurred'));
  });
});
