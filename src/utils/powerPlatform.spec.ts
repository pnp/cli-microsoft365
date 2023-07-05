import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import auth from '../Auth.js';
import { powerPlatform } from './powerPlatform.js';
import { sinonUtil } from "./sinonUtil.js";
import { formatting } from './formatting.js';

const validChatbotName = 'CLI 365 Chatbot';
const envUrl = 'https://contoso-dev.api.crm4.dynamics.com';
const chatbotResponse = {
  value: [
    {
      authenticationtrigger: 0,
      '_owningbusinessunit_value': '6da087c1-1c4d-ed11-bba1-000d3a2caf7f',
      statuscode: 1,
      createdon: '2022-11-19T10:42:22Z',
      statecode: 0,
      schemaname: 'new_bot_23f5f58697fd43d595eb451c9797a53d',
      '_ownerid_value': '5fa787c1-1c4d-ed11-bba1-000d3a2caf7f',
      overwritetime: '1900-01-01T00:00:00Z',
      name: validChatbotName,
      solutionid: 'fd140aae-4df4-11dd-bd17-0019b9312238',
      ismanaged: false,
      versionnumber: 1421457,
      language: 1033,
      '_modifiedby_value': '5f91d7a7-5f46-494a-80fa-5c18b0221351',
      '_modifiedonbehalfby_value': '5fa787c1-1c4d-ed11-bba1-000d3a2caf7f',
      modifiedon: '2022-11-19T10:42:24Z',
      componentstate: 0,
      botid: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893',
      '_createdby_value': '5fa787c1-1c4d-ed11-bba1-000d3a2caf7f',
      componentidunique: 'cdcd6496-e25d-4ad1-91cf-3f4d547fdd23',
      authenticationmode: 1,
      '_owninguser_value': '5fa787c1-1c4d-ed11-bba1-000d3a2caf7f',
      accesscontrolpolicy: 0,
      runtimeprovider: 0,
      '_publishedby_value': 'John Doe',
      authenticationconfiguration: null,
      authorizedsecuritygroupids: null,
      overriddencreatedon: null,
      applicationmanifestinformation: null,
      importsequencenumber: null,
      synchronizationstatus: null,
      template: null,
      '_providerconnectionreferenceid_value': null,
      configuration: null,
      utcconversiontimezonecode: null,
      publishedon: '2022-11-19T10:43:24Z',
      '_createdonbehalfby_value': null,
      iconbase64: null,
      supportedlanguages: null,
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

  it('throws error when multiple chatbotss with same name were found', async () => {
    const multipleChatbotResponse = {
      value: [
        { ["botid"]: '69703efe-4149-ed11-bba2-000d3adf7537' },
        { ["botid"]: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' }
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(validChatbotName)}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return multipleChatbotResponse;
        }
      }

      throw 'Invalid request';
    });

    try {
      await powerPlatform.getChatbotByName(envUrl, validChatbotName);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Multiple chatbots with name '${validChatbotName}' found: 69703efe-4149-ed11-bba2-000d3adf7537,3a081d91-5ea8-40a7-8ac9-abbaa3fcb893`));
    }
  });

  it('throws error when no chatbot found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(validChatbotName)}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ "value": [] });
        }
      }

      throw 'Invalid request';
    });

    try {
      await powerPlatform.getChatbotByName(envUrl, validChatbotName);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The specified chatbot '${validChatbotName}' does not exist.`));
    }
  });

  it('retrieves a specific chatbot with the name parameter', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(validChatbotName)}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return chatbotResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    const actual = await powerPlatform.getChatbotByName(envUrl, validChatbotName);
    assert.strictEqual(actual, chatbotResponse.value[0]);
  });
});