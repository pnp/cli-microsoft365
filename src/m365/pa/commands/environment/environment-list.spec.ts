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

describe(commands.ENVIRONMENT_LIST, () => {
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
    assert.strictEqual(command.name, commands.ENVIRONMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'displayName']);
  });

  it('retrieves Microsoft App environments (debug)', async () => {
    const env: any = { value: [{ "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-2172edf9f447", "name": "Default-2ca3eaa5-140f-4175-9563-2172edf9f447", "location": "europe", "type": "Microsoft.PowerApps/environments", "properties": { "azureRegionHint": "westeurope", "displayName": "Contoso (default) (contoso)", "createdTime": "2016-10-28T10:32:54.1945519Z", "createdBy": { "id": "SYSTEM", "displayName": "SYSTEM", "type": "NotSpecified" }, "lastModifiedTime": "2020-07-28T08:58:12.5785779Z", "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Administrator", "email": "administrator@contoso.nl", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-2172edf9f447", "userPrincipalName": "administrator@contoso.nl" }, "provisioningState": "Succeeded", "creationType": "DefaultTenant", "environmentSku": "Default", "environmentType": "Production", "isDefault": true, "runtimeEndpoints": { "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com", "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com", "microsoft.PowerApps": "https://europe.api.powerapps.com", "microsoft.Flow": "https://europe.api.flow.microsoft.com", "microsoft.PowerAppsAdvisor": "https://europe.api.advisor.powerapps.com", "microsoft.ApiManagement": "https://management.EUR.azure-apihub.net" }, "linkedEnvironmentMetadata": { "type": "Dynamics365Instance", "resourceId": "6d590664-6f39-41f4-9e8b-e95bc6bb1f1f", "friendlyName": "Contoso (default)", "uniqueName": "org185e622f", "domainName": "contoso", "version": "9.2.20122.00144", "instanceUrl": "https://contoso.crm4.dynamics.com/", "instanceApiUrl": "https://contoso.api.crm4.dynamics.com", "baseLanguage": 1033, "instanceState": "Ready", "createdTime": "2018-11-14T10:17:47.033Z", "modifiedTime": "2021-01-27T18:59:34.0883701Z", "hostNameSuffix": "crm4.dynamics.com", "bapSolutionId": "00000001-0000-0000-0001-00000000009b", "creationTemplates": ["D365_CDS"], "webApiVersion": "v9.0" }, "retentionPeriod": "P7D", "lifecycleAuthority": "Environment", "states": { "management": { "id": "NotSpecified" }, "runtime": { "id": "Enabled" } }, "updateCadence": { "id": "Frequent" }, "connectedGroups": [], "protectionStatus": { "keyManagedBy": "Microsoft" } } }] };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.PowerApps/environments?api-version=2017-08-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return env;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(env.value));
  });

  it('retrieves Microsoft App environments', async () => {
    const env: any = { value: [{ "id": "/providers/Microsoft.PowerApps/environments/Default-2ca3eaa5-140f-4175-9563-2172edf9f447", "name": "Default-2ca3eaa5-140f-4175-9563-2172edf9f447", "location": "europe", "type": "Microsoft.PowerApps/environments", "properties": { "azureRegionHint": "westeurope", "displayName": "Contoso (default) (contoso)", "createdTime": "2016-10-28T10:32:54.1945519Z", "createdBy": { "id": "SYSTEM", "displayName": "SYSTEM", "type": "NotSpecified" }, "lastModifiedTime": "2020-07-28T08:58:12.5785779Z", "lastModifiedBy": { "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e", "displayName": "Administrator", "email": "administrator@contoso.nl", "type": "User", "tenantId": "2ca3eaa5-140f-4175-9563-2172edf9f447", "userPrincipalName": "administrator@contoso.nl" }, "provisioningState": "Succeeded", "creationType": "DefaultTenant", "environmentSku": "Default", "environmentType": "Production", "isDefault": true, "runtimeEndpoints": { "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com", "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com", "microsoft.PowerApps": "https://europe.api.powerapps.com", "microsoft.Flow": "https://europe.api.flow.microsoft.com", "microsoft.PowerAppsAdvisor": "https://europe.api.advisor.powerapps.com", "microsoft.ApiManagement": "https://management.EUR.azure-apihub.net" }, "linkedEnvironmentMetadata": { "type": "Dynamics365Instance", "resourceId": "6d590664-6f39-41f4-9e8b-e95bc6bb1f1f", "friendlyName": "Contoso (default)", "uniqueName": "org185e622f", "domainName": "contoso", "version": "9.2.20122.00144", "instanceUrl": "https://contoso.crm4.dynamics.com/", "instanceApiUrl": "https://contoso.api.crm4.dynamics.com", "baseLanguage": 1033, "instanceState": "Ready", "createdTime": "2018-11-14T10:17:47.033Z", "modifiedTime": "2021-01-27T18:59:34.0883701Z", "hostNameSuffix": "crm4.dynamics.com", "bapSolutionId": "00000001-0000-0000-0001-00000000009b", "creationTemplates": ["D365_CDS"], "webApiVersion": "v9.0" }, "retentionPeriod": "P7D", "lifecycleAuthority": "Environment", "states": { "management": { "id": "NotSpecified" }, "runtime": { "id": "Enabled" } }, "updateCadence": { "id": "Frequent" }, "connectedGroups": [], "protectionStatus": { "keyManagedBy": "Microsoft" } } }] };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.PowerApps/environments?api-version=2017-08-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return env;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(env.value));
  });

  it('correctly handles no environments', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.PowerApps/environments?api-version=2017-08-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { value: [] };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly([]));
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

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
