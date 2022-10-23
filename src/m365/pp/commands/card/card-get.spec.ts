import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./card-get');

describe(commands.CARD_GET, () => {
  //#region Mocked Responses
  const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };
  const cardResponse: any = {
    "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238",
    "modifiedon": "2022-10-11T08:52:12Z",
    "_owninguser_value": "7d48edd3-69fd-ec11-82e5-000d3ab87733",
    "overriddencreatedon": null,
    "ismanaged": false,
    "schemaversion": null,
    "tags": null,
    "importsequencenumber": null,
    "componentidunique": "d7c1acb5-37a4-4873-b24e-34b18c15c6a5",
    "_modifiedonbehalfby_value": null,
    "componentstate": 0,
    "statecode": 0,
    "name": "DummyCard",
    "versionnumber": 3044006,
    "utcconversiontimezonecode": null,
    "cardid": "69703efe-4149-ed11-bba2-000d3adf7537",
    "publishdate": null,
    "_createdonbehalfby_value": null,
    "_modifiedby_value": "7d48edd3-69fd-ec11-82e5-000d3ab87733",
    "createdon": "2022-10-11T08:52:12Z",
    "overwritetime": "1900-01-01T00:00:00Z",
    "_owningbusinessunit_value": "2199f44c-195b-ec11-8f8f-000d3adca49c",
    "hiddentags": null,
    "description": " ",
    "appdefinition": "{\"screens\":{\"main\":{\"template\":{\"type\":\"AdaptiveCard\",\"body\":[{\"type\":\"TextBlock\",\"size\":\"Medium\",\"weight\":\"bolder\",\"text\":\"Your card title goes here\"},{\"type\":\"TextBlock\",\"text\":\"Add and remove element to customize your new card.\",\"wrap\":true}],\"actions\":[],\"$schema\":\"http://adaptivecards.io/schemas/1.4.0/adaptive-card.json\",\"version\":\"1.4\"},\"verbs\":{\"submit\":\"echo\"}}},\"sampleData\":{\"main\":{}},\"connections\":{},\"variables\":{},\"flows\":{}}",
    "statuscode": 1,
    "remixsourceid": null,
    "sizes": null,
    "_owningteam_value": null,
    "coowners": null,
    "_createdby_value": "7d48edd3-69fd-ec11-82e5-000d3ab87733",
    "_ownerid_value": "7d48edd3-69fd-ec11-82e5-000d3ab87733",
    "publishsourceid": null,
    "timezoneruleversionnumber": null,
    "iscustomizable": {
      "Value": true,
      "CanBeChanged": true,
      "ManagedPropertyLogicalName": "iscustomizableanddeletable"
    },
    "owninguser": {
      "azureactivedirectoryobjectid": "88e85b64-e687-4e0b-bbf4-f42f5f8e574c",
      "fullname": "Contoso Admin",
      "systemuserid": "7d48edd3-69fd-ec11-82e5-000d3ab87733",
      "ownerid": "7d48edd3-69fd-ec11-82e5-000d3ab87733"
    }
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CARD_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'cardid', 'publishdate', 'createdon', 'modifiedon']);
  });

  it('retrieves a specific card', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards(3a081d91-5ea8-40a7-8ac9-abbaa3fcb893)`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return cardResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', id: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' } });
    assert(loggerLogSpy.calledWith(cardResponse));

  });

  it('retrieves cards as admin', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards(3a081d91-5ea8-40a7-8ac9-abbaa3fcb893)`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return cardResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', id: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893', asAdmin: true } });
    assert(loggerLogSpy.calledWith(cardResponse));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      throw `Resource '' does not exist or one of its queried reference-property objects are not present`;
    });

    await assert.rejects(command.action(logger, { options: { debug: false, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', id: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' } }), new CommandError("Resource '' does not exist or one of its queried reference-property objects are not present"));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});