import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { accessToken } from '../../../../utils/accessToken.js';
import commands from '../../commands.js';
import command from './environment-get.js';

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
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
      request.get,
      pid.getProcessName,
      session.getId
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

  it('fails validation when no options specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when name is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: environmentName
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when default is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      default: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when only asAdmin is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      asAdmin: true
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when both name and asAdmin are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: environmentName,
      asAdmin: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when both default and asAdmin are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      default: true,
      asAdmin: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when both name and default are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: environmentName,
      default: true
    });
    assert.strictEqual(actual.success, false);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = `Resource '' does not exist or one of its queried reference-property objects are not present`;
    sinon.stub(request, 'get').callsFake(async () => {
      throw errorMessage;
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        name: environmentName
      })
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
      options: commandOptionsSchema.parse({
        name: environmentName,
        verbose: true
      })
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
      options: commandOptionsSchema.parse({
        default: true,
        verbose: true
      })
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
      options: commandOptionsSchema.parse({
        name: environmentName,
        asAdmin: true,
        verbose: true
      })
    });

    assert(loggerLogSpy.calledWith(environmentResponse));
  });
});
