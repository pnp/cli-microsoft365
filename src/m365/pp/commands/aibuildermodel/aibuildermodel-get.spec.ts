import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './aibuildermodel-get.js';
import { session } from '../../../../utils/session.js';
import { settingsNames } from '../../../../settingsNames.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.AIBUILDERMODEL_GET, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validName = 'CLI 365 AI Builder Model';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const aiBuilderModelResponse = {
    "value": [
      {
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
      }
    ]
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.AIBUILDERMODEL_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, name: validName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws error when multiple AI builder models with same name were found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

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

    await assert.rejects(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        name: validName
      }
    }), new CommandError("Multiple AI builder models with name 'CLI 365 AI Builder Model' found. Found: 69703efe-4149-ed11-bba2-000d3adf7537, 3a081d91-5ea8-40a7-8ac9-abbaa3fcb893."));
  });

  it('handles selecting single result when multiple AI builder models with the specified name found and cli is set to prompt', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

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

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(aiBuilderModelResponse.value[0]);

    await command.action(logger, { options: { verbose: true, environment: validEnvironment, name: validName } });
    assert(loggerLogSpy.calledWith(aiBuilderModelResponse.value[0]));
  });

  it('throws error when no AI builder model found', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${validName}' and iscustomizable/Value eq true`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ "value": [] });
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        name: validName
      }
    }), new CommandError(`The specified AI builder model '${validName}' does not exist.`));
  });

  it('retrieves a specific AI builder model with the name parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${validName}' and iscustomizable/Value eq true`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return aiBuilderModelResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironment, name: validName } });
    assert(loggerLogSpy.calledWith(aiBuilderModelResponse.value[0]));
  });

  it('retrieves a specific AI builder model with the id parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels(${validId})?$filter=iscustomizable/Value eq true`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return aiBuilderModelResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironment, id: validId } });
    assert(loggerLogSpy.calledWith(aiBuilderModelResponse.value[0]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${validName}' and iscustomizable/Value eq true`)) {
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

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironment, name: validName } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});