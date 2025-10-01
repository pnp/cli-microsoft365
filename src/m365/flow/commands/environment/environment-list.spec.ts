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
import command from './environment-list.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';

describe(commands.ENVIRONMENT_LIST, () => {
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
    assert.strictEqual(command.name, commands.ENVIRONMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'displayName']);
  });

  it('retrieves Microsoft Flow environments (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            value: [
              {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (default)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": true,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              },
              {
                "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (test)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": false,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              }
            ]
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) });
    assert(loggerLogSpy.calledWith([
      {
        "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "location": "europe",
        "type": "Microsoft.ProcessSimple/environments",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "properties": {
          "displayName": "Contoso (default)",
          "createdTime": "2018-03-22T20:20:46.08653Z",
          "createdBy": {
            "id": "SYSTEM",
            "displayName": "SYSTEM",
            "type": "NotSpecified"
          },
          "provisioningState": "Succeeded",
          "creationType": "DefaultTenant",
          "environmentSku": "Default",
          "environmentType": "Production",
          "isDefault": true,
          "azureRegionHint": "westeurope",
          "runtimeEndpoints": {
            "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
            "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
            "microsoft.PowerApps": "https://europe.api.powerapps.com",
            "microsoft.Flow": "https://europe.api.flow.microsoft.com"
          }
        },
        "displayName": "Contoso (default)"
      },
      {
        "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "location": "europe",
        "type": "Microsoft.ProcessSimple/environments",
        "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "properties": {
          "displayName": "Contoso (test)",
          "createdTime": "2018-03-22T20:20:46.08653Z",
          "createdBy": {
            "id": "SYSTEM",
            "displayName": "SYSTEM",
            "type": "NotSpecified"
          },
          "provisioningState": "Succeeded",
          "creationType": "DefaultTenant",
          "environmentSku": "Default",
          "environmentType": "Production",
          "isDefault": false,
          "azureRegionHint": "westeurope",
          "runtimeEndpoints": {
            "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
            "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
            "microsoft.PowerApps": "https://europe.api.powerapps.com",
            "microsoft.Flow": "https://europe.api.flow.microsoft.com"
          }
        },
        "displayName": "Contoso (test)"
      }
    ]));
  });

  it('retrieves Microsoft Flow environments', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            value: [
              {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (default)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": true,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              },
              {
                "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (test)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": false,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              }
            ]
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.safeParse({}) });
    assert(loggerLogSpy.calledWith([
      {
        "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "location": "europe",
        "type": "Microsoft.ProcessSimple/environments",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "properties": {
          "displayName": "Contoso (default)",
          "createdTime": "2018-03-22T20:20:46.08653Z",
          "createdBy": {
            "id": "SYSTEM",
            "displayName": "SYSTEM",
            "type": "NotSpecified"
          },
          "provisioningState": "Succeeded",
          "creationType": "DefaultTenant",
          "environmentSku": "Default",
          "environmentType": "Production",
          "isDefault": true,
          "azureRegionHint": "westeurope",
          "runtimeEndpoints": {
            "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
            "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
            "microsoft.PowerApps": "https://europe.api.powerapps.com",
            "microsoft.Flow": "https://europe.api.flow.microsoft.com"
          }
        },
        "displayName": "Contoso (default)"
      },
      {
        "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "location": "europe",
        "type": "Microsoft.ProcessSimple/environments",
        "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "properties": {
          "displayName": "Contoso (test)",
          "createdTime": "2018-03-22T20:20:46.08653Z",
          "createdBy": {
            "id": "SYSTEM",
            "displayName": "SYSTEM",
            "type": "NotSpecified"
          },
          "provisioningState": "Succeeded",
          "creationType": "DefaultTenant",
          "environmentSku": "Default",
          "environmentType": "Production",
          "isDefault": false,
          "azureRegionHint": "westeurope",
          "runtimeEndpoints": {
            "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
            "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
            "microsoft.PowerApps": "https://europe.api.powerapps.com",
            "microsoft.Flow": "https://europe.api.flow.microsoft.com"
          }
        },
        "displayName": "Contoso (test)"
      }
    ]));
  });

  it('retrieves Microsoft Flow environments with output text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            value: [
              {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (default)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": true,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              },
              {
                "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (test)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": false,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              }
            ]
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ output: 'text' }) });
    assert(loggerLogSpy.calledWith([
      {
        "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "location": "europe",
        "type": "Microsoft.ProcessSimple/environments",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "properties": {
          "displayName": "Contoso (default)",
          "createdTime": "2018-03-22T20:20:46.08653Z",
          "createdBy": {
            "id": "SYSTEM",
            "displayName": "SYSTEM",
            "type": "NotSpecified"
          },
          "provisioningState": "Succeeded",
          "creationType": "DefaultTenant",
          "environmentSku": "Default",
          "environmentType": "Production",
          "isDefault": true,
          "azureRegionHint": "westeurope",
          "runtimeEndpoints": {
            "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
            "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
            "microsoft.PowerApps": "https://europe.api.powerapps.com",
            "microsoft.Flow": "https://europe.api.flow.microsoft.com"
          }
        },
        "displayName": "Contoso (default)"
      },
      {
        "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "location": "europe",
        "type": "Microsoft.ProcessSimple/environments",
        "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "properties": {
          "displayName": "Contoso (test)",
          "createdTime": "2018-03-22T20:20:46.08653Z",
          "createdBy": {
            "id": "SYSTEM",
            "displayName": "SYSTEM",
            "type": "NotSpecified"
          },
          "provisioningState": "Succeeded",
          "creationType": "DefaultTenant",
          "environmentSku": "Default",
          "environmentType": "Production",
          "isDefault": false,
          "azureRegionHint": "westeurope",
          "runtimeEndpoints": {
            "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
            "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
            "microsoft.PowerApps": "https://europe.api.powerapps.com",
            "microsoft.Flow": "https://europe.api.flow.microsoft.com"
          }
        },
        "displayName": "Contoso (test)"
      }
    ]));
  });

  it('retrieves Microsoft Flow environments with output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            value: [
              {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (default)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": true,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              },
              {
                "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "location": "europe",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "properties": {
                  "displayName": "Contoso (test)",
                  "createdTime": "2018-03-22T20:20:46.08653Z",
                  "createdBy": {
                    "id": "SYSTEM",
                    "displayName": "SYSTEM",
                    "type": "NotSpecified"
                  },
                  "provisioningState": "Succeeded",
                  "creationType": "DefaultTenant",
                  "environmentSku": "Default",
                  "environmentType": "Production",
                  "isDefault": false,
                  "azureRegionHint": "westeurope",
                  "runtimeEndpoints": {
                    "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
                    "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
                    "microsoft.PowerApps": "https://europe.api.powerapps.com",
                    "microsoft.Flow": "https://europe.api.flow.microsoft.com"
                  }
                }
              }
            ]
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ output: 'json' }) });
    assert(loggerLogSpy.calledWithExactly([
      {
        "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "location": "europe",
        "type": "Microsoft.ProcessSimple/environments",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "properties": {
          "displayName": "Contoso (default)",
          "createdTime": "2018-03-22T20:20:46.08653Z",
          "createdBy": {
            "id": "SYSTEM",
            "displayName": "SYSTEM",
            "type": "NotSpecified"
          },
          "provisioningState": "Succeeded",
          "creationType": "DefaultTenant",
          "environmentSku": "Default",
          "environmentType": "Production",
          "isDefault": true,
          "azureRegionHint": "westeurope",
          "runtimeEndpoints": {
            "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
            "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
            "microsoft.PowerApps": "https://europe.api.powerapps.com",
            "microsoft.Flow": "https://europe.api.flow.microsoft.com"
          }
        }
      },
      {
        "name": "Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "location": "europe",
        "type": "Microsoft.ProcessSimple/environments",
        "id": "/providers/Microsoft.ProcessSimple/environments/Test-d87a7535-dd31-4437-bfe1-95340acd55c5",
        "properties": {
          "displayName": "Contoso (test)",
          "createdTime": "2018-03-22T20:20:46.08653Z",
          "createdBy": {
            "id": "SYSTEM",
            "displayName": "SYSTEM",
            "type": "NotSpecified"
          },
          "provisioningState": "Succeeded",
          "creationType": "DefaultTenant",
          "environmentSku": "Default",
          "environmentType": "Production",
          "isDefault": false,
          "azureRegionHint": "westeurope",
          "runtimeEndpoints": {
            "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
            "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
            "microsoft.PowerApps": "https://europe.api.powerapps.com",
            "microsoft.Flow": "https://europe.api.flow.microsoft.com"
          }
        }
      }
    ]));
  });

  it('correctly handles no environments', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { value: [] };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.safeParse({}) });
    assert(loggerLogSpy.calledWithExactly([]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: `Resource '' does not exist or one of its queried reference-property objects are not present`
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.safeParse({}) } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
