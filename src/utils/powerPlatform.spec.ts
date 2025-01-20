import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import auth from '../Auth.js';
import { powerPlatform } from './powerPlatform.js';
import { sinonUtil } from "./sinonUtil.js";
import { cli } from '../cli/cli.js';
import { settingsNames } from '../settingsNames.js';

const environment = 'Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0';
const powerPageResponse = {
  value: [{
    "@odata.metadata": "https://api.powerplatform.com/powerpages/environments/Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0/websites/$metadata#Websites",
    "id": "4916bb2c-91e1-4716-91d5-b6171928fac9",
    "name": "Site 1",
    "createdOn": "2024-10-27T12:00:03",
    "templateName": "DefaultPortalTemplate",
    "websiteUrl": "https://site-0uaq9.powerappsportals.com",
    "tenantId": "727dc1e9-3cd1-4d1f-8102-ab5c936e52f0",
    "dataverseInstanceUrl": "https://org0cd4b2b9.crm4.dynamics.com/",
    "environmentName": "Contoso (default)",
    "environmentId": "Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0",
    "dataverseOrganizationId": "2d58aeac-74d4-4939-98d1-e05a70a655ba",
    "selectedBaseLanguage": 1033,
    "customHostNames": [],
    "websiteRecordId": "5eb107a6-5ac2-4e1c-a3b9-d5c21bbc10ce",
    "subdomain": "site-0uaq9",
    "packageInstallStatus": "Installed",
    "type": "Trial",
    "trialExpiringInDays": 86,
    "suspendedWebsiteDeletingInDays": 93,
    "packageVersion": "9.6.9.39",
    "isEarlyUpgradeEnabled": false,
    "isCustomErrorEnabled": true,
    "applicationUserAadAppId": "3f57aca7-5051-41b2-989d-26da8af7a53e",
    "ownerId": "33469a62-c3af-4cfe-b893-854eceab96da",
    "status": "OperationComplete",
    "siteVisibility": "private",
    "dataModel": "Enhanced"
  },
  {
    "@odata.metadata": "https://api.powerplatform.com/powerpages/environments/Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0/websites/$metadata#Websites",
    "id": "dc2b0aa4-4449-4667-b1a8-41017b8f874c",
    "name": "Site 2",
    "createdOn": "2024-10-27T12:02:59",
    "templateName": "DefaultPortalTemplate",
    "websiteUrl": "https://site-aa9wk.powerappsportals.com",
    "tenantId": "727dc1e9-3cd1-4d1f-8102-ab5c936e52f0",
    "dataverseInstanceUrl": "https://org0cd4b2b9.crm4.dynamics.com/",
    "environmentName": "Contoso (default)",
    "environmentId": "Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0",
    "dataverseOrganizationId": "2d58aeac-74d4-4939-98d1-e05a70a655ba",
    "selectedBaseLanguage": 1033,
    "customHostNames": [],
    "websiteRecordId": "bc59fb78-d685-4b70-b9e3-531ece45536d",
    "subdomain": "site-aa9wk",
    "packageInstallStatus": "Installed",
    "type": "Trial",
    "trialExpiringInDays": 86,
    "suspendedWebsiteDeletingInDays": 93,
    "packageVersion": "9.6.9.39",
    "isEarlyUpgradeEnabled": false,
    "isCustomErrorEnabled": true,
    "applicationUserAadAppId": "3f57aca7-5051-41b2-989d-26da8af7a53e",
    "ownerId": "33469a62-c3af-4cfe-b893-854eceab96da",
    "status": "OperationComplete",
    "siteVisibility": "private",
    "dataModel": "Enhanced"
  }]
};

const MultiplePowerPageResponseWithSameName = {
  ...powerPageResponse,
  value: [
    ...powerPageResponse.value.slice(0, 1),
    {
      ...powerPageResponse.value[1],
      name: "Site 1"
    }
  ]
};

describe('utils/powerPlatform', () => {
  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    auth.connection.active = true;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => settingName === settingsNames.prompt ? false : defaultValue);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      powerPlatform.getWebsiteById,
      powerPlatform.getWebsiteByName,
      powerPlatform.getWebsiteByUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('returns correct dynamics url as admin', async () => {
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      throw 'Invalid request';
    });

    const actual = await powerPlatform.getDynamicsInstanceApiUrl('someRandomGuid', true);
    assert.strictEqual(actual, 'https://contoso-dev.api.crm4.dynamics.com');
  });

  it('returns correct dynamics url', async () => {
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      throw 'Invalid request';
    });

    const actual = await powerPlatform.getDynamicsInstanceApiUrl('someRandomGuid', false);
    assert.strictEqual(actual, 'https://contoso-dev.api.crm4.dynamics.com');
  });

  it('handles no environment found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        throw Error('Random Error');
      }

      return 'Invalid request';
    });

    try {
      await powerPlatform.getDynamicsInstanceApiUrl('someRandomGuid', false);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The environment 'someRandomGuid' could not be retrieved. See the inner exception for more details: Random Error`));
    }
  });

  //#region Power Page websites
  it('returns correct Power Page website by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.powerplatform.com/powerpages/environments/${environment}/websites/4916bb2c-91e1-4716-91d5-b6171928fac9?api-version=2022-03-01-preview`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return powerPageResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    const actual = await powerPlatform.getWebsiteById(environment, '4916bb2c-91e1-4716-91d5-b6171928fac9');
    assert.strictEqual(actual, powerPageResponse.value[0]);
  });

  it('handles error when using id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.powerplatform.com/powerpages/environments/${environment}/websites/be13f9af-f73d-48d6-99c0-7097c03282fc?api-version=2022-03-01-preview`)) {
        throw Error('Random Error');
      }

      return 'Invalid request';
    });

    try {
      await powerPlatform.getWebsiteById(environment, 'be13f9af-f73d-48d6-99c0-7097c03282fc');
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The specified Power Page website with id 'be13f9af-f73d-48d6-99c0-7097c03282fc' does not exist.`));
    }
  });

  it('returns correct Power Page website by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return powerPageResponse;
        }
      }

      throw 'Invalid request';
    });

    const actual = await powerPlatform.getWebsiteByName(environment, 'Site 1');
    assert.strictEqual(actual, powerPageResponse.value[0]);
  });

  it('throws error message when multiple Power Page websites were found using name', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`) {
        return MultiplePowerPageResponseWithSameName;
      }

      throw 'Invalid Request';
    });

    await assert.rejects(powerPlatform.getWebsiteByName(environment, 'Site 1'),
      new Error(`Multiple Power Page websites with name 'Site 1' found Found: https://site-0uaq9.powerappsportals.com, https://site-aa9wk.powerappsportals.com.`));
  });

  it('handles selecting single result when multiple Power Page websites with the specified name found using name and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`) {
        return MultiplePowerPageResponseWithSameName;
      }

      throw 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(MultiplePowerPageResponseWithSameName.value[0]);

    const actual = await powerPlatform.getWebsiteByName(environment, 'Site 1');
    assert.deepStrictEqual(actual, MultiplePowerPageResponseWithSameName.value[0]);
  });

  it('handles no Power Page website found when using name', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`)) {
        return { value: [] };
      }

      return 'Invalid request';
    });

    await assert.rejects(powerPlatform.getWebsiteByName(environment, 'Site 1'), Error(`The specified Power Page website 'Site 1' does not exist.`));
  });

  it('returns correct Power Page website by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return powerPageResponse;
        }
      }

      throw 'Invalid request';
    });

    const actual = await powerPlatform.getWebsiteByUrl(environment, 'https://site-0uaq9.powerappsportals.com');
    assert.strictEqual(actual, powerPageResponse.value[0]);
  });

  it('handles no Power Page website found when using url', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`)) {
        return { value: [] };
      }

      return 'Invalid request';
    });

    await assert.rejects(powerPlatform.getWebsiteByUrl(environment, 'https://site-0uaq9.powerappsportals.com'), Error(`The specified Power Page website with url 'https://site-0uaq9.powerappsportals.com' does not exist.`));
  });
  // #endregion
});