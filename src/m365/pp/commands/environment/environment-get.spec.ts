import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./environment-get');

describe(commands.ENVIRONMENT_GET, () => {
  const environmentName = 'Default-de347bc8-1aeb-4406-8cb3-97db021cadb4';
  const environmentResponse = {
    "id": `/providers/Microsoft.BusinessAppPlatform/environments/Default-de347bc8-1aeb-4406-8cb3-97db021cadb4`,
    "type": "Microsoft.BusinessAppPlatform/environments",
    "location": "unitedstates",
    "name": "Default-de347bc8-1aeb-4406-8cb3-97db021cadb4",
    "properties": {
      "displayName": "contoso (default)",
      "isDefault": true
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
      request.get,
      pid.getProcessName,
      session.getId
    ]);
  });

  after(() => {
    sinon.restore();
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

  it('correctly handles API OData error', async () => {
    const errorMessage = `Resource '' does not exist or one of its queried reference-property objects are not present`;
    sinon.stub(request, 'get').callsFake(async () => {
      throw errorMessage;
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: environmentName
      }
    }), new CommandError(errorMessage));
  });

  it('retrieves Microsoft Power Platform environment by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(environmentName)}?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return environmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: environmentName,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith(environmentResponse));
  });

  it('retrieves default Microsoft Power Platform environment', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/~Default?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return environmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith(environmentResponse));
  });

  it('retrieves Microsoft Power Platform environment as Admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return environmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: environmentName,
        asAdmin: true,
        verbose: true
      }
    });

    assert(loggerLogSpy.calledWith(environmentResponse));
  });
});
