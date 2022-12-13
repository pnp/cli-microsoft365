import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./chatbot-get');

describe(commands.CHATBOT_GET, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validName = 'CLI 365 Chatbot';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const botResponse = {
    "value": [
      {
        "authenticationtrigger": 0,
        "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f",
        "statuscode": 1,
        "createdon": "2022-11-19T10:42:22Z",
        "statecode": 0,
        "schemaname": "new_bot_23f5f58697fd43d595eb451c9797a53d",
        "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "overwritetime": "1900-01-01T00:00:00Z",
        "name": "CLI 365 Chatbot",
        "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238",
        "ismanaged": false,
        "versionnumber": 1421457,
        "language": 1033,
        "_modifiedby_value": "5f91d7a7-5f46-494a-80fa-5c18b0221351",
        "_modifiedonbehalfby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "modifiedon": "2022-11-19T10:42:24Z",
        "componentstate": 0,
        "botid": "3a081d91-5ea8-40a7-8ac9-abbaa3fcb893",
        "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "componentidunique": "cdcd6496-e25d-4ad1-91cf-3f4d547fdd23",
        "authenticationmode": 1,
        "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "accesscontrolpolicy": 0,
        "runtimeprovider": 0,
        "_publishedby_value": "John Doe",
        "authenticationconfiguration": null,
        "authorizedsecuritygroupids": null,
        "overriddencreatedon": null,
        "applicationmanifestinformation": null,
        "importsequencenumber": null,
        "synchronizationstatus": null,
        "template": null,
        "_providerconnectionreferenceid_value": null,
        "configuration": null,
        "utcconversiontimezonecode": null,
        "publishedon": "2022-11-19T10:43:24Z",
        "_createdonbehalfby_value": null,
        "iconbase64": null,
        "supportedlanguages": null,
        "_owningteam_value": null,
        "timezoneruleversionnumber": null,
        "iscustomizable": {
          "Value": true,
          "CanBeChanged": true,
          "ManagedPropertyLogicalName": "iscustomizableanddeletable"
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
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHATBOT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'botid', 'publishedon', 'createdon', 'modifiedon']);
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

  it('throws error when multiple chatbots found with the same name', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    const multipleBotsResponse = {
      value: [
        { botid: '69703efe-4149-ed11-bba2-000d3adf7537' },
        { botid: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' }
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(validName)}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return multipleBotsResponse;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        environment: validEnvironment,
        name: validName
      }
    }), new CommandError(`Multiple chatbots with name '${validName}' found: ${multipleBotsResponse.value.map(x => x.botid).join(',')}`));
  });

  it('throws error when no chatbot with name was found', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(validName)}'`)) {
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
    }), new CommandError(`The specified chatbot '${validName}' does not exist.`));
  });

  it('retrieves a specific chatbot with the name parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(validName)}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return botResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, environment: validEnvironment, name: validName } });
    assert(loggerLogSpy.calledWith(botResponse.value[0]));
  });

  it('retrieves a specific chatbot with the id parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots(${validId})`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return botResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environment: validEnvironment, id: validId } });
    assert(loggerLogSpy.calledWith(botResponse.value[0]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(validName)}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          throw {
            error: {
              error: {
                message: `bot With Id = ${validId} Does Not Exist`
              }
            }
          };
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { debug: false, environment: validEnvironment, name: validName } } as any),
      new CommandError(`bot With Id = ${validId} Does Not Exist`));
  });
});
