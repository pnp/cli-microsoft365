import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./environment-list');

describe(commands.ENVIRONMENT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ENVIRONMENT_LIST), true);
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
          return Promise.resolve(env);
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true } });
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
          return Promise.resolve(env);
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: {} });
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
          return Promise.resolve(env);
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { asAdmin: true } });
    assert(loggerLogSpy.calledWith(env.value));
  });
  it('correctly handles no environments', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.BusinessAppPlatform/environments?api-version=2020-10-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({
            value: []
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
