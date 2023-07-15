import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import auth from '../Auth.js';
import { powerPlatform } from './powerPlatform.js';
import { sinonUtil } from "./sinonUtil.js";
import { Logger } from '../cli/Logger.js';

const validName = 'CLI 365 AI Builder Model';
const envUrl = 'https://contoso-dev.api.crm4.dynamics.com';
const aiBuilderModelResponse = {
  value: [
    {
      statecode: 0,
      '_msdyn_templateid_value': '10707e4e-1d56-e911-8194-000d3a6cd5a5',
      'msdyn_modelcreationcontext': '{}',
      createdon: '2022-11-29T11:58:45Z',
      '_ownerid_value': '5fa787c1-1c4d-ed11-bba1-000d3a2caf7f',
      modifiedon: '2022-11-29T11:58:45Z',
      'msdyn_sharewithorganizationoncreate': false,
      'msdyn_aimodelidunique': 'b0328b67-47e2-4202-8189-e617ec9a88bd',
      solutionid: 'fd140aae-4df4-11dd-bd17-0019b9312238',
      ismanaged: false,
      versionnumber: 1458121,
      'msdyn_name': 'Document Processing 11/29/2022, 12:58:43 PM',
      introducedversion: '1.0',
      statuscode: 0,
      '_modifiedby_value': '5fa787c1-1c4d-ed11-bba1-000d3a2caf7f',
      overwritetime: '1900-01-01T00:00:00Z',
      componentstate: 0,
      '_createdby_value': '5fa787c1-1c4d-ed11-bba1-000d3a2caf7f',
      '_owningbusinessunit_value': '6da087c1-1c4d-ed11-bba1-000d3a2caf7f',
      '_owninguser_value': '5fa787c1-1c4d-ed11-bba1-000d3a2caf7f',
      'msdyn_aimodelid': '08ffffbe-ec1c-4e64-b64b-dd1db926c613',
      '_msdyn_activerunconfigurationid_value': null,
      overriddencreatedon: null,
      '_msdyn_retrainworkflowid_value': null,
      importsequencenumber: null,
      '_msdyn_scheduleinferenceworkflowid_value': null,
      '_modifiedonbehalfby_value': null,
      utcconversiontimezonecode: null,
      '_createdonbehalfby_value': null,
      '_owningteam_value': null,
      timezoneruleversionnumber: null,
      iscustomizable: {
        Value: true,
        CanBeChanged: true,
        ManagedPropertyLogicalName: 'iscustomizableanddeletable'
      }
    }
  ]
};

describe('utils/powerPlatform', () => {
  let logger: Logger;
  let log: string[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
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

  it('throws error when multiple AI builder models with same name were found', async () => {
    const multipleAiBuilderModelsResponse = {
      value: [
        { ["msdyn_aimodelid"]: '69703efe-4149-ed11-bba2-000d3adf7537' },
        { ["msdyn_aimodelid"]: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' }
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${validName}' and iscustomizable/Value eq true`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return multipleAiBuilderModelsResponse;
        }
      }

      throw 'Invalid request';
    });

    try {
      await powerPlatform.getAiBuilderModelByName(envUrl, validName);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Multiple AI builder models with name 'CLI 365 AI Builder Model' found: 69703efe-4149-ed11-bba2-000d3adf7537,3a081d91-5ea8-40a7-8ac9-abbaa3fcb893`));
    }
  });

  it('throws error when no AI builder model found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${validName}' and iscustomizable/Value eq true`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ "value": [] });
        }
      }

      throw 'Invalid request';
    });

    try {
      await powerPlatform.getAiBuilderModelByName(envUrl, validName);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The specified AI builder model 'CLI 365 AI Builder Model' does not exist.`));
    }
  });

  it('retrieves a specific AI builder model with the name parameter', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${validName}' and iscustomizable/Value eq true`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return aiBuilderModelResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    const actual = await powerPlatform.getAiBuilderModelByName(envUrl, validName, logger, true);
    assert.strictEqual(actual, aiBuilderModelResponse.value[0]);
  });
});