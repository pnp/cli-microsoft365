import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import auth from '../Auth.js';
import { powerPlatform } from './powerPlatform.js';
import { sinonUtil } from "./sinonUtil.js";

const validSolutionPublisherName = 'CLI 365 Solution';
const envUrl = 'https://contoso-dev.api.crm4.dynamics.com';
const solutionPublisherResponse = {
  value: [
    {
      publisherid: "d21aab70-79e7-11dd-8874-00188b01e34f",
      uniquename: validSolutionPublisherName,
      friendlyname: validSolutionPublisherName,
      versionnumber: 1226559,
      isreadonly: false,
      customizationprefix: "",
      customizationoptionvalueprefix: 0
    }
  ]
};

describe('utils/powerPlatform', () => {
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('returns correct dynamics url as admin', async () => {
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }

      return Promise.reject('Invalid request');
    });

    const actual = await powerPlatform.getDynamicsInstanceApiUrl('someRandomGuid', true);
    assert.strictEqual(actual, 'https://contoso-dev.api.crm4.dynamics.com');
  });

  it('returns correct dynamics url', async () => {
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }

      return Promise.reject('Invalid request');
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

  it('throws error when multiple solution publishers with same name were found', async () => {
    const multipleSolutionPublishersResponse = {
      value: [
        { ["publisherid"]: '69703efe-4149-ed11-bba2-000d3adf7537' },
        { ["publisherid"]: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' }
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$filter=friendlyname eq \'${validSolutionPublisherName}\'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return multipleSolutionPublishersResponse;
        }
      }

      throw 'Invalid request';
    });

    try {
      await powerPlatform.getSolutionPublisherByName(envUrl, validSolutionPublisherName);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Multiple solution publishers with name '${validSolutionPublisherName}' found: 69703efe-4149-ed11-bba2-000d3adf7537,3a081d91-5ea8-40a7-8ac9-abbaa3fcb893`));
    }
  });

  it('throws error when no solution publisher found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$filter=friendlyname eq \'${validSolutionPublisherName}\'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ "value": [] });
        }
      }

      throw 'Invalid request';
    });

    try {
      await powerPlatform.getSolutionPublisherByName(envUrl, validSolutionPublisherName);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The specified solution publisher '${validSolutionPublisherName}' does not exist.`));
    }
  });

  it('retrieves a specific solution publisher with the name parameter', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$filter=friendlyname eq \'${validSolutionPublisherName}\'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return solutionPublisherResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    const actual = await powerPlatform.getSolutionPublisherByName(envUrl, validSolutionPublisherName);
    assert.strictEqual(actual, solutionPublisherResponse.value[0]);
  });
});