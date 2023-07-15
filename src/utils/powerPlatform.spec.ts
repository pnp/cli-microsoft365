import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import auth from '../Auth.js';
import { powerPlatform } from './powerPlatform.js';
import { sinonUtil } from "./sinonUtil.js";
import { Logger } from '../cli/Logger.js';

const validCardName = 'CLI 365 Card';
const envUrl = 'https://contoso-dev.api.crm4.dynamics.com';
const cardResponse = {
  value: [
    {
      solutionid: 'fd140aae-4df4-11dd-bd17-0019b9312238',
      modifiedon: '2022-10-11T08:52:12Z',
      '_owninguser_value': '7d48edd3-69fd-ec11-82e5-000d3ab87733',
      overriddencreatedon: null,
      ismanaged: false,
      schemaversion: null,
      tags: null,
      importsequencenumber: null,
      componentidunique: 'd7c1acb5-37a4-4873-b24e-34b18c15c6a5',
      '_modifiedonbehalfby_value': null,
      componentstate: 0,
      statecode: 0,
      name: 'DummyCard',
      versionnumber: 3044006,
      utcconversiontimezonecode: null,
      cardid: '69703efe-4149-ed11-bba2-000d3adf7537',
      publishdate: null,
      '_createdonbehalfby_value': null,
      '_modifiedby_value': '7d48edd3-69fd-ec11-82e5-000d3ab87733',
      createdon: '2022-10-11T08:52:12Z',
      overwritetime: '1900-01-01T00:00:00Z',
      '_owningbusinessunit_value': '2199f44c-195b-ec11-8f8f-000d3adca49c',
      hiddentags: null,
      description: ' ',
      appdefinition: '{\'screens\':{\'main\':{\'template\':{\'type\':\'AdaptiveCard\',\'body\':[{\'type\':\'TextBlock\',\'size\':\'Medium\',\'weight\':\'bolder\',\'text\':\'Your card title goes here\'},{\'type\':\'TextBlock\',\'text\':\'Add and remove element to customize your new card.\',\'wrap\':true}],\'actions\':[],\'$schema\':\'http://adaptivecards.io/schemas/1.4.0/adaptive-card.json\',\'version\':\'1.4\'},\'verbs\':{\'submit\':\'echo\'}}},\'sampleData\':{\'main\':{}},\'connections\':{},\'variables\':{},\'flows\':{}}',
      statuscode: 1,
      remixsourceid: null,
      sizes: null,
      '_owningteam_value': null,
      coowners: null,
      '_createdby_value': '7d48edd3-69fd-ec11-82e5-000d3ab87733',
      '_ownerid_value': '7d48edd3-69fd-ec11-82e5-000d3ab87733',
      publishsourceid: null,
      timezoneruleversionnumber: null,
      iscustomizable: {
        Value: true,
        CanBeChanged: true,
        ManagedPropertyLogicalName: 'iscustomizableanddeletable'
      },
      owninguser: {
        azureactivedirectoryobjectid: '88e85b64-e687-4e0b-bbf4-f42f5f8e574c',
        fullname: 'Contoso Admin',
        systemuserid: '7d48edd3-69fd-ec11-82e5-000d3ab87733',
        ownerid: '7d48edd3-69fd-ec11-82e5-000d3ab87733'
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

  it('throws error when multiple cards with same name were found', async () => {
    const multipleCardsResponse = {
      value: [
        { ["cardid"]: '69703efe-4149-ed11-bba2-000d3adf7537' },
        { ["cardid"]: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' }
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards?$filter=name eq '${validCardName}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return multipleCardsResponse;
        }
      }

      throw 'Invalid request';
    });

    try {
      await powerPlatform.getCardByName(envUrl, validCardName, logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Multiple cards with name 'CLI 365 Card' found: 69703efe-4149-ed11-bba2-000d3adf7537,3a081d91-5ea8-40a7-8ac9-abbaa3fcb893`));
    }
  });

  it('throws error when no card found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards?$filter=name eq '${validCardName}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ "value": [] });
        }
      }

      throw 'Invalid request';
    });

    try {
      await powerPlatform.getCardByName(envUrl, validCardName);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The specified card 'CLI 365 Card' does not exist.`));
    }
  });

  it('retrieves a specific card with the name parameter', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards?$filter=name eq '${validCardName}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return cardResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    const actual = await powerPlatform.getCardByName(envUrl, validCardName);
    assert.strictEqual(actual, cardResponse.value[0]);
  });
});