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
import * as PpAiBuilderModelGetCommand from './aibuildermodel-get';
const command: Command = require('./aibuildermodel-remove');

describe(commands.AIBUILDERMODEL_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '08ffffbe-ec1c-4e64-b64b-dd1db926c613';
  const validName = 'CLI 365 Ai Builder Model';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const aiBuilderModelResponse = `{
    "statecode": 0,
    "_msdyn_templateid_value": "10707e4e-1d56-e911-8194-000d3a6cd5a5",
    "msdyn_modelcreationcontext": "{}",
    "createdon": "2022-11-29T11:58:45Z",
    "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
    "modifiedon": "2022-11-29T11:58:45Z",
    "msdyn_sharewithorganizationoncreate": false,
    "msdyn_aimodelidunique": "b0328b67-47e2-4202-8189-e617ec9a88bd",
    "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238",
    "ismanaged": false,
    "versionnumber": 1458121,
    "msdyn_name": "Document Processing 11/29/2022, 12:58:43 PM",
    "introducedversion": "1.0",
    "statuscode": 0,
    "_modifiedby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
    "overwritetime": "1900-01-01T00:00:00Z",
    "componentstate": 0,
    "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
    "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f",
    "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
    "msdyn_aimodelid": "08ffffbe-ec1c-4e64-b64b-dd1db926c613",
    "_msdyn_activerunconfigurationid_value": null,
    "overriddencreatedon": null,
    "_msdyn_retrainworkflowid_value": null,
    "importsequencenumber": null,
    "_msdyn_scheduleinferenceworkflowid_value": null,
    "_modifiedonbehalfby_value": null,
    "utcconversiontimezonecode": null,
    "_createdonbehalfby_value": null,
    "_owningteam_value": null,
    "timezoneruleversionnumber": null,
    "iscustomizable": {
      "Value": true,
      "CanBeChanged": true,
      "ManagedPropertyLogicalName": "iscustomizableanddeletable"
    }
  }`;
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
      request.delete,
      powerPlatform.getDynamicsInstanceApiUrl,
      Cli.prompt,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.AIBUILDERMODEL_REMOVE);
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

  it('prompts before removing the specified AI builder model owned by the currently signed-in user when confirm option not passed', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

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

  it('aborts removing the specified AI builder model owned by the currently signed-in user when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        environment: validEnvironment,
        id: validId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified AI builder model owned by the currently signed-in user when prompt confirmed', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === PpAiBuilderModelGetCommand) {
        return ({
          stdout: aiBuilderModelResponse
        });
      }

      throw new CommandError('Unknown case');
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels(${validId})`) {
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
        debug: true,
        environment: validEnvironment,
        name: validName
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified AI builder model without confirmation prompt', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels(${validId})`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
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

    sinon.stub(request, 'delete').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

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
