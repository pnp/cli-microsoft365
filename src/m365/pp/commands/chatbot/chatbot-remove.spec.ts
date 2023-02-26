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
import * as PpChatbotGetCommand from './chatbot-get';
const command: Command = require('./chatbot-remove');

describe(commands.CHATBOT_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validName = 'CLI 365 Chatbot';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      powerPlatform.getDynamicsInstanceApiUrl,
      Cli.prompt,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name, commands.CHATBOT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('prompts before removing the specified chatbot owned by the currently signed-in user when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        environment: validEnvironment,
        id: validId
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified chatbot owned by the currently signed-in user when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        environment: validEnvironment,
        id: validId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified chatbot owned by the currently signed-in user when prompt confirmed by name', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === PpChatbotGetCommand) {
        return ({
          stdout: `{ "authenticationtrigger": 0, "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f", "statuscode": 1, "createdon": "2022-11-19T10:42:22Z", "statecode": 0, "schemaname": "new_bot_23f5f58697fd43d595eb451c9797a53d", "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f", "name": "CLI 365 Chatbot", "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238", "ismanaged": false, "versionnumber": 1429641, "publishedon": "2022-11-19T19:19:53Z", "timezoneruleversionnumber": 0, "language": 1033, "_modifiedby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f", "overwritetime": "1900-01-01T00:00:00Z", "modifiedon": "2022-11-19T20:19:57Z", "componentstate": 0, "botid": "3a081d91-5ea8-40a7-8ac9-abbaa3fcb893", "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f", "componentidunique": "cdcd6496-e25d-4ad1-91cf-3f4d547fdd23", "authenticationmode": 1, "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f", "accesscontrolpolicy": 0, "runtimeprovider": 0, "_publishedby_value": null, "authenticationconfiguration": null, "authorizedsecuritygroupids": null, "overriddencreatedon": null, "applicationmanifestinformation": null, "importsequencenumber": null, "synchronizationstatus": null, "_modifiedonbehalfby_value": null, "template": null, "_providerconnectionreferenceid_value": null, "configuration": null, "utcconversiontimezonecode": null, "_createdonbehalfby_value": null, "iconbase64": null, "supportedlanguages": null, "_owningteam_value": null, "iscustomizable": { "Value": true, "CanBeChanged": true, "ManagedPropertyLogicalName": "iscustomizableanddeletable" } }`
        });
      }

      throw new CommandError('Unknown case');
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots(${validId})/Microsoft.Dynamics.CRM.PvaDeleteBot?tag=deprovisionbotondelete`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, {
      options: {
        verbose: true,
        environment: validEnvironment,
        name: validName
      }
    });
    assert(postStub.called);
  });

  it('removes the specified chatbot without confirmation prompt by id', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots(${validId})/Microsoft.Dynamics.CRM.PvaDeleteBot?tag=deprovisionbotondelete`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        environment: validEnvironment,
        id: validId,
        confirm: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'post').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environment: validEnvironment,
        id: validId,
        confirm: true
      }
    }), new CommandError(errorMessage));
  });
});
