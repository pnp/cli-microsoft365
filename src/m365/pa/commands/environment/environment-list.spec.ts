import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
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
      pid.getProcessName
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
    const env: any = { value: [{ "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-2172edf9f447", "name": "Default-2ca3eaa5-140f-4175-9563-2172edf9f447", "location": "europe", "type": "Microsoft.PowerApps/environments", "properties": { "azureRegionHint": "westeurope", "displayName": "Contoso (default) (contoso)", "createdTime": "2016-10-28T10:32:54.1945519Z", "createdBy": { "id": "SYSTEM", "displayName": "SYSTEM", "type": "NotSpecified" }, "lastModifiedTime": "2020-07-28T08:58:12.5785779Z", "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Administrator", "email": "administrator@contoso.nl", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-2172edf9f447", "userPrincipalName": "administrator@contoso.nl" }, "provisioningState": "Succeeded", "creationType": "DefaultTenant", "environmentSku": "Default", "environmentType": "Production", "isDefault": true, "runtimeEndpoints": { "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com", "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com", "microsoft.PowerApps": "https://europe.api.powerapps.com", "microsoft.Flow": "https://europe.api.flow.microsoft.com", "microsoft.PowerAppsAdvisor": "https://europe.api.advisor.powerapps.com", "microsoft.ApiManagement": "https://management.EUR.azure-apihub.net" }, "linkedEnvironmentMetadata": { "type": "Dynamics365Instance", "resourceId": "6d590664-6f39-41f4-9e8b-e95bc6bb1f1f", "friendlyName": "Contoso (default)", "uniqueName": "org185e622f", "domainName": "contoso", "version": "9.2.20122.00144", "instanceUrl": "https://contoso.crm4.dynamics.com/", "instanceApiUrl": "https://contoso.api.crm4.dynamics.com", "baseLanguage": 1033, "instanceState": "Ready", "createdTime": "2018-11-14T10:17:47.033Z", "modifiedTime": "2021-01-27T18:59:34.0883701Z", "hostNameSuffix": "crm4.dynamics.com", "bapSolutionId": "00000001-0000-0000-0001-00000000009b", "creationTemplates": ["D365_CDS"], "webApiVersion": "v9.0" }, "retentionPeriod": "P7D", "lifecycleAuthority": "Environment", "states": { "management": { "id": "NotSpecified" }, "runtime": { "id": "Enabled" } }, "updateCadence": { "id": "Frequent" }, "connectedGroups": [], "protectionStatus": { "keyManagedBy": "Microsoft" } } }] };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.PowerApps/environments?api-version=2017-08-01`) > -1) {
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

  it('retrieves Microsoft App environments', async () => {
    const env: any = { value: [{ "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-2172edf9f447", "name": "Default-2ca3eaa5-140f-4175-9563-2172edf9f447", "location": "europe", "type": "Microsoft.PowerApps/environments", "properties": { "azureRegionHint": "westeurope", "displayName": "Contoso (default) (contoso)", "createdTime": "2016-10-28T10:32:54.1945519Z", "createdBy": { "id": "SYSTEM", "displayName": "SYSTEM", "type": "NotSpecified" }, "lastModifiedTime": "2020-07-28T08:58:12.5785779Z", "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Administrator", "email": "administrator@contoso.nl", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-2172edf9f447", "userPrincipalName": "administrator@contoso.nl" }, "provisioningState": "Succeeded", "creationType": "DefaultTenant", "environmentSku": "Default", "environmentType": "Production", "isDefault": true, "runtimeEndpoints": { "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com", "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com", "microsoft.PowerApps": "https://europe.api.powerapps.com", "microsoft.Flow": "https://europe.api.flow.microsoft.com", "microsoft.PowerAppsAdvisor": "https://europe.api.advisor.powerapps.com", "microsoft.ApiManagement": "https://management.EUR.azure-apihub.net" }, "linkedEnvironmentMetadata": { "type": "Dynamics365Instance", "resourceId": "6d590664-6f39-41f4-9e8b-e95bc6bb1f1f", "friendlyName": "Contoso (default)", "uniqueName": "org185e622f", "domainName": "contoso", "version": "9.2.20122.00144", "instanceUrl": "https://contoso.crm4.dynamics.com/", "instanceApiUrl": "https://contoso.api.crm4.dynamics.com", "baseLanguage": 1033, "instanceState": "Ready", "createdTime": "2018-11-14T10:17:47.033Z", "modifiedTime": "2021-01-27T18:59:34.0883701Z", "hostNameSuffix": "crm4.dynamics.com", "bapSolutionId": "00000001-0000-0000-0001-00000000009b", "creationTemplates": ["D365_CDS"], "webApiVersion": "v9.0" }, "retentionPeriod": "P7D", "lifecycleAuthority": "Environment", "states": { "management": { "id": "NotSpecified" }, "runtime": { "id": "Enabled" } }, "updateCadence": { "id": "Frequent" }, "connectedGroups": [], "protectionStatus": { "keyManagedBy": "Microsoft" } } }] };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.PowerApps/environments?api-version=2017-08-01`) > -1) {
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

  it('correctly handles no environments', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.PowerApps/environments?api-version=2017-08-01`) > -1) {
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
