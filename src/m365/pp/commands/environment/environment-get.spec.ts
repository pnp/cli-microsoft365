import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./environment-get');

describe(commands.ENVIRONMENT_GET, () => {
  const name = '5ca1c616-6060-46ba-abc1-18d312f1cb3a';
  const nameEnvironmentResponse: any = {
    "value": [
      {
        "id": `/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${name}`,
        "type": "Microsoft.BusinessAppPlatform/scopes/environments",
        "location": "unitedstates",
        "name": `${name}`,
        "properties": {
          "displayName": "My Power Platform Environment",
          "isDefault": false
        }
      }
    ]
  };

  const defaultEnvironmentResponse: any = {
    "value": [
      {
        "id": `/providers/Microsoft.BusinessAppPlatform/environments/Default-de347bc8-1aeb-4406-8cb3-97db021cadb4`,
        "type": "Microsoft.BusinessAppPlatform/environments",
        "location": "unitedstates",
        "name": "Default-de347bc8-1aeb-4406-8cb3-97db021cadb4",
        "properties": {
          "displayName": "contoso (default)",
          "isDefault": true
        }
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
      telemetry.trackEvent
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
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'id']);
  });

  it('correctly handles no environments', async () => {
    const errorMessage = 'The specified Power Platform environment does not exist';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            value: []
          };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: name
      }
    }), new CommandError(errorMessage));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = `Resource '' does not exist or one of its queried reference-property objects are not present`;
    sinon.stub(request, 'get').callsFake(async () => {
      throw errorMessage;
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: name
      }
    }), new CommandError(errorMessage));
  });

  it('retrieves Microsoft Power Platform environment by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return nameEnvironmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: name
      }
    });
    assert(loggerLogSpy.calledWith(nameEnvironmentResponse.value[0]));
  });

  it('retrieves default Microsoft Power Platform environment', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return defaultEnvironmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(defaultEnvironmentResponse.value[0]));
  });

  it('retrieves Microsoft Power Platform environment as Admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return nameEnvironmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        asAdmin: true,
        name: name
      }
    });

    assert(loggerLogSpy.calledWith(nameEnvironmentResponse.value[0]));
  });
});
