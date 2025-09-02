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
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './environment-get.js';
import { FlowEnvironmentDetails } from './FlowEnvironmentDetails.js';

describe(commands.ENVIRONMENT_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  const flowResponse: FlowEnvironmentDetails = {
    name: "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
    location: "europe",
    type: "Microsoft.ProcessSimple/environments",
    id: "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
    properties: {
      displayName: "Contoso (default)",
      createdTime: "2018-03-22T20:20:46.08653Z",
      createdBy: {
        id: "SYSTEM",
        displayName: "SYSTEM",
        type: "NotSpecified"
      },
      provisioningState: "Succeeded",
      creationType: "DefaultTenant",
      environmentSku: "Default",
      environmentType: "Production",
      isDefault: true,
      azureRegionHint: "westeurope",
      runtimeEndpoints: {
        "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
        "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
        "microsoft.PowerApps": "https://europe.api.powerapps.com",
        "microsoft.Flow": "https://europe.api.flow.microsoft.com"
      }
    }
  };

  const flowResponseText: FlowEnvironmentDetails = {
    displayName: "Contoso (default)",
    provisioningState: "Succeeded",
    environmentSku: "Default",
    azureRegionHint: "westeurope",
    isDefault: true,
    name: "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
    location: "europe",
    type: "Microsoft.ProcessSimple/environments",
    id: "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
    properties: {
      displayName: "Contoso (default)",
      createdTime: "2018-03-22T20:20:46.08653Z",
      createdBy: {
        id: "SYSTEM",
        displayName: "SYSTEM",
        type: "NotSpecified"
      },
      provisioningState: "Succeeded",
      creationType: "DefaultTenant",
      environmentSku: "Default",
      environmentType: "Production",
      isDefault: true,
      azureRegionHint: "westeurope",
      runtimeEndpoints: {
        "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
        "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
        "microsoft.PowerApps": "https://europe.api.powerapps.com",
        "microsoft.Flow": "https://europe.api.flow.microsoft.com"
      }
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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

  it('fails validation when no options specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when name is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when default is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      default: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when both name and default are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5',
      default: true
    });
    assert.strictEqual(actual.success, false);
  });

  it('retrieves information about the specified environment (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5?api-version=2016-11-01`)) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'json',
        debug: true,
        name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5'
      })
    });
    assert(loggerLogSpy.calledWith(flowResponse));
  });

  it('retrieves information about the specified environment', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5?api-version=2016-11-01`)) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'json',
        name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5'
      })
    });
    assert(loggerLogSpy.calledWith(flowResponse));
  });

  it('retrieves information about the specified environment with output text', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5?api-version=2016-11-01`)) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'text',
        name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5'
      })
    });
    assert(loggerLogSpy.calledWith(flowResponseText));
  });

  it('retrieves information about the default environment', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/~default?api-version=2016-11-01`)) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        output: 'json',
        verbose: true,
        default: true
      })
    });
    assert(loggerLogSpy.calledWith(flowResponse));
  });

  it('correctly handles no environment found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
      }
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6'
      })
    }), new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`));
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

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5'
      })
    }), new CommandError('An error has occurred'));
  });
});
