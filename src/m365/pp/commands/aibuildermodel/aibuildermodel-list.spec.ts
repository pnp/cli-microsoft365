import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { powerPlatform } from '../../../../utils/powerPlatform';
const command: Command = require('./aibuildermodel-list');

describe(commands.AIBUILDERMODEL_LIST, () => {
  //#region Mocked Responses
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const validEnvironment = "4be50206-9576-4237-8b17-38d8aadfaa36";
  const modelsResponse: any = {
    "value": [
      {
        "@odata.etag": "W/\"1458121\"",
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
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.AIBUILDERMODEL_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['msdyn_name', 'msdyn_aimodelid', 'createdon', 'modifiedon']);
  });

  it('retrieves AI Builder models', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/msdyn_aimodels?$filter=iscustomizable/Value eq true`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return modelsResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environment: validEnvironment } });
    assert(loggerLogSpy.calledWith(modelsResponse.value));

  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/msdyn_aimodels?$filter=iscustomizable/Value eq true`) {
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

    await assert.rejects(command.action(logger, { options: { environment: validEnvironment } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
