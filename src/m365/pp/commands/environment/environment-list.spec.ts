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
import { accessToken } from '../../../../utils/accessToken.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import { z } from 'zod';

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

  it('retrieves Microsoft App environments (debug)', async () => {
    const env: any = {
      "value": [
        {
          "id": "/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a",
          "type": "Microsoft.BusinessAppPlatform/scopes/environments",
          "location": "unitedstates",
          "name": "5ca1c616-6060-46ba-abc1-18d312f1cb3a",
          "properties": {
            "azureRegion": "westus",
            "displayName": "My Power Platform Environment",
            "description": "This is my environment purpose description",
            "createdTime": "2020-10-22T04:38:17.8550157Z",
            "createdBy": {
              "id": "0f747967-84c4-4f29-84c2-682fb00390c8",
              "displayName": "ServicePrincipal",
              "type": "ServicePrincipal",
              "tenantId": "5ca1c616-6060-46ba-abc1-18d312f1cb3a"
            },
            "lastModifiedTime": "2021-02-22T18:38:08.4718532Z",
            "provisioningState": "Succeeded",
            "creationType": "User",
            "environmentSku": "Sandbox",
            "isDefault": false,
            "capacity": [
              {
                "capacityType": "Database",
                "actualConsumption": 1392.68,
                "ratedConsumption": 1392.68,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              },
              {
                "capacityType": "File",
                "actualConsumption": 1567.697,
                "ratedConsumption": 1567.697,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              },
              {
                "capacityType": "Log",
                "actualConsumption": 0,
                "ratedConsumption": 0,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              }
            ],
            "addons": [
              {
                "addonType": "AppPass",
                "allocated": 1,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PerFlowPlan",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PortalViews",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PortalLogins",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "AI",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "AppPassForTeams",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PAUnattendedRPA",
                "allocated": 0,
                "addonUnit": "Unit"
              }
            ],
            "clientUris": {
              "admin": "https://admin.powerplatform.microsoft.com/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a/hub",
              "maker": "https://make.powerapps.com/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a/home"
            },
            "runtimeEndpoints": {
              "microsoft.BusinessAppPlatform": "https://unitedstates.api.bap.microsoft.com",
              "microsoft.CommonDataModel": "https://unitedstates.api.cds.microsoft.com",
              "microsoft.PowerApps": "https://unitedstates.api.powerapps.com",
              "microsoft.Flow": "https://unitedstates.api.flow.microsoft.com",
              "microsoft.PowerAppsAdvisor": "https://unitedstates.api.advisor.powerapps.com",
              "microsoft.ApiManagement": "https://management.usa.azure-apihub.net"
            },
            "databaseType": "CommonDataService",
            "linkedEnvironmentMetadata": {
              "resourceId": "3b48b422-0b37-4070-8054-601867eb8b23",
              "friendlyName": "My Power Platform Environment",
              "uniqueName": "96c28a9e98934bf798bb71c9d92134",
              "domainName": "org0fadb1dd",
              "version": "9.2.21013.00152",
              "instanceUrl": "https://org0fadb1dd.crm.dynamics.com/",
              "instanceApiUrl": "https://org0fadb1dd.api.crm.dynamics.com",
              "baseLanguage": 1033,
              "instanceState": "Ready",
              "createdTime": "2020-10-22T04:38:24.003Z",
              "backgroundOperationsState": "Enabled",
              "scaleGroup": "NAMCRMLIVESG644",
              "platformSku": "Standard"
            },
            "notificationMetadata": {
              "state": "NotSpecified",
              "branding": "NotSpecific"
            },
            "retentionPeriod": "P7D",
            "states": {
              "management": {
                "id": "Ready"
              },
              "runtime": {
                "id": "Enabled"
              }
            },
            "updateCadence": {
              "id": "Frequent"
            },
            "retentionDetails": {
              "retentionPeriod": "P7D",
              "backupsAvailableFromDateTime": "2021-02-16T05:42:52.2822636Z"
            },
            "protectionStatus": {
              "keyManagedBy": "Microsoft"
            },
            "cluster": {
              "number": "108"
            },
            "connectedGroups": []
          }
        }
      ]
    };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.BusinessAppPlatform/environments?api-version=2020-10-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return env;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) });
    assert(loggerLogSpy.calledWith(env.value));
  });

  it('retrieves Microsoft Power Platform environments', async () => {
    const env: any = {
      "value": [
        {
          "id": "/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a",
          "type": "Microsoft.BusinessAppPlatform/scopes/environments",
          "location": "unitedstates",
          "name": "5ca1c616-6060-46ba-abc1-18d312f1cb3a",
          "properties": {
            "azureRegion": "westus",
            "displayName": "My Power Platform Environment",
            "description": "This is my environment purpose description",
            "createdTime": "2020-10-22T04:38:17.8550157Z",
            "createdBy": {
              "id": "0f747967-84c4-4f29-84c2-682fb00390c8",
              "displayName": "ServicePrincipal",
              "type": "ServicePrincipal",
              "tenantId": "5ca1c616-6060-46ba-abc1-18d312f1cb3a"
            },
            "lastModifiedTime": "2021-02-22T18:38:08.4718532Z",
            "provisioningState": "Succeeded",
            "creationType": "User",
            "environmentSku": "Sandbox",
            "isDefault": false,
            "capacity": [
              {
                "capacityType": "Database",
                "actualConsumption": 1392.68,
                "ratedConsumption": 1392.68,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              },
              {
                "capacityType": "File",
                "actualConsumption": 1567.697,
                "ratedConsumption": 1567.697,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              },
              {
                "capacityType": "Log",
                "actualConsumption": 0,
                "ratedConsumption": 0,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              }
            ],
            "addons": [
              {
                "addonType": "AppPass",
                "allocated": 1,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PerFlowPlan",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PortalViews",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PortalLogins",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "AI",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "AppPassForTeams",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PAUnattendedRPA",
                "allocated": 0,
                "addonUnit": "Unit"
              }
            ],
            "clientUris": {
              "admin": "https://admin.powerplatform.microsoft.com/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a/hub",
              "maker": "https://make.powerapps.com/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a/home"
            },
            "runtimeEndpoints": {
              "microsoft.BusinessAppPlatform": "https://unitedstates.api.bap.microsoft.com",
              "microsoft.CommonDataModel": "https://unitedstates.api.cds.microsoft.com",
              "microsoft.PowerApps": "https://unitedstates.api.powerapps.com",
              "microsoft.Flow": "https://unitedstates.api.flow.microsoft.com",
              "microsoft.PowerAppsAdvisor": "https://unitedstates.api.advisor.powerapps.com",
              "microsoft.ApiManagement": "https://management.usa.azure-apihub.net"
            },
            "databaseType": "CommonDataService",
            "linkedEnvironmentMetadata": {
              "resourceId": "3b48b422-0b37-4070-8054-601867eb8b23",
              "friendlyName": "My Power Platform Environment",
              "uniqueName": "96c28a9e98934bf798bb71c9d92134",
              "domainName": "org0fadb1dd",
              "version": "9.2.21013.00152",
              "instanceUrl": "https://org0fadb1dd.crm.dynamics.com/",
              "instanceApiUrl": "https://org0fadb1dd.api.crm.dynamics.com",
              "baseLanguage": 1033,
              "instanceState": "Ready",
              "createdTime": "2020-10-22T04:38:24.003Z",
              "backgroundOperationsState": "Enabled",
              "scaleGroup": "NAMCRMLIVESG644",
              "platformSku": "Standard"
            },
            "notificationMetadata": {
              "state": "NotSpecified",
              "branding": "NotSpecific"
            },
            "retentionPeriod": "P7D",
            "states": {
              "management": {
                "id": "Ready"
              },
              "runtime": {
                "id": "Enabled"
              }
            },
            "updateCadence": {
              "id": "Frequent"
            },
            "retentionDetails": {
              "retentionPeriod": "P7D",
              "backupsAvailableFromDateTime": "2021-02-16T05:42:52.2822636Z"
            },
            "protectionStatus": {
              "keyManagedBy": "Microsoft"
            },
            "cluster": {
              "number": "108"
            },
            "connectedGroups": []
          }
        }
      ]
    };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.BusinessAppPlatform/environments?api-version=2020-10-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return env;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(loggerLogSpy.calledWith(env.value));
  });

  it('retrieves Microsoft Power Platform environments as Admin', async () => {
    const env: any = {
      "value": [
        {
          "id": "/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a",
          "type": "Microsoft.BusinessAppPlatform/scopes/environments",
          "location": "unitedstates",
          "name": "5ca1c616-6060-46ba-abc1-18d312f1cb3a",
          "properties": {
            "azureRegion": "westus",
            "displayName": "My Power Platform Environment",
            "description": "This is my environment purpose description",
            "createdTime": "2020-10-22T04:38:17.8550157Z",
            "createdBy": {
              "id": "0f747967-84c4-4f29-84c2-682fb00390c8",
              "displayName": "ServicePrincipal",
              "type": "ServicePrincipal",
              "tenantId": "5ca1c616-6060-46ba-abc1-18d312f1cb3a"
            },
            "lastModifiedTime": "2021-02-22T18:38:08.4718532Z",
            "provisioningState": "Succeeded",
            "creationType": "User",
            "environmentSku": "Sandbox",
            "isDefault": false,
            "capacity": [
              {
                "capacityType": "Database",
                "actualConsumption": 1392.68,
                "ratedConsumption": 1392.68,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              },
              {
                "capacityType": "File",
                "actualConsumption": 1567.697,
                "ratedConsumption": 1567.697,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              },
              {
                "capacityType": "Log",
                "actualConsumption": 0,
                "ratedConsumption": 0,
                "capacityUnit": "MB",
                "updatedOn": "2021-02-23T04:41:01Z"
              }
            ],
            "addons": [
              {
                "addonType": "AppPass",
                "allocated": 1,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PerFlowPlan",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PortalViews",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PortalLogins",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "AI",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "AppPassForTeams",
                "allocated": 0,
                "addonUnit": "Unit"
              },
              {
                "addonType": "PAUnattendedRPA",
                "allocated": 0,
                "addonUnit": "Unit"
              }
            ],
            "clientUris": {
              "admin": "https://admin.powerplatform.microsoft.com/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a/hub",
              "maker": "https://make.powerapps.com/environments/5ca1c616-6060-46ba-abc1-18d312f1cb3a/home"
            },
            "runtimeEndpoints": {
              "microsoft.BusinessAppPlatform": "https://unitedstates.api.bap.microsoft.com",
              "microsoft.CommonDataModel": "https://unitedstates.api.cds.microsoft.com",
              "microsoft.PowerApps": "https://unitedstates.api.powerapps.com",
              "microsoft.Flow": "https://unitedstates.api.flow.microsoft.com",
              "microsoft.PowerAppsAdvisor": "https://unitedstates.api.advisor.powerapps.com",
              "microsoft.ApiManagement": "https://management.usa.azure-apihub.net"
            },
            "databaseType": "CommonDataService",
            "linkedEnvironmentMetadata": {
              "resourceId": "3b48b422-0b37-4070-8054-601867eb8b23",
              "friendlyName": "My Power Platform Environment",
              "uniqueName": "96c28a9e98934bf798bb71c9d92134",
              "domainName": "org0fadb1dd",
              "version": "9.2.21013.00152",
              "instanceUrl": "https://org0fadb1dd.crm.dynamics.com/",
              "instanceApiUrl": "https://org0fadb1dd.api.crm.dynamics.com",
              "baseLanguage": 1033,
              "instanceState": "Ready",
              "createdTime": "2020-10-22T04:38:24.003Z",
              "backgroundOperationsState": "Enabled",
              "scaleGroup": "NAMCRMLIVESG644",
              "platformSku": "Standard"
            },
            "notificationMetadata": {
              "state": "NotSpecified",
              "branding": "NotSpecific"
            },
            "retentionPeriod": "P7D",
            "states": {
              "management": {
                "id": "Ready"
              },
              "runtime": {
                "id": "Enabled"
              }
            },
            "updateCadence": {
              "id": "Frequent"
            },
            "retentionDetails": {
              "retentionPeriod": "P7D",
              "backupsAvailableFromDateTime": "2021-02-16T05:42:52.2822636Z"
            },
            "protectionStatus": {
              "keyManagedBy": "Microsoft"
            },
            "cluster": {
              "number": "108"
            },
            "connectedGroups": []
          }
        }
      ]
    };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2020-10-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return env;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ asAdmin: true }) });
    assert(loggerLogSpy.calledWith(env.value));
  });

  it('correctly handles no environments', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.BusinessAppPlatform/environments?api-version=2020-10-01`) > -1) {
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

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw {
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
