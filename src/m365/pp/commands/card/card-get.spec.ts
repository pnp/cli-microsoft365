import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./card-get');

describe(commands.CARD_GET, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validName = 'CLI 365 Card';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const cardResponse = {
    "value": [
      {
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
      }
    ]
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl
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
    assert.strictEqual(command.name.startsWith(commands.CARD_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'cardid', 'publishdate', 'createdon', 'modifiedon']);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environment: validEnvironment,
        id: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { environment: validEnvironment, id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { environment: validEnvironment, name: validName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws error when no card found', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards?$filter=name eq '${validName}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ "value": [] });
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        environment: validEnvironment,
        name: validName
      }
    }), new CommandError(`The specified card '${validName}' does not exist.`));
  });

  it('throws error when multiple cards with same name were found', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    const multipleCardsResponse = {
      value: [
        { cardid: '69703efe-4149-ed11-bba2-000d3adf7537' },
        { cardid: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' }
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards?$filter=name eq '${validName}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return multipleCardsResponse;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        environment: validEnvironment,
        name: validName
      }
    }), new CommandError(`Multiple cards with name '${validName}' found: ${multipleCardsResponse.value.map(x => x.cardid).join(',')}`));
  });

  it('retrieves a specific card with the name parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards?$filter=name eq '${validName}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return cardResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, environment: validEnvironment, name: validName } });
    assert(loggerLogSpy.calledWith(cardResponse.value[0]));
  });

  it('retrieves a specific card with the id parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards(${validId})`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return cardResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environment: validEnvironment, id: validId } });
    assert(loggerLogSpy.calledWith(cardResponse.value[0]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards?$filter=name eq '${validName}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
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
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environment: validEnvironment, name: validName } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
