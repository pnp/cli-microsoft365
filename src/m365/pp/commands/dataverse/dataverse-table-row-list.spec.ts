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
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './dataverse-table-row-list.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.DATAVERSE_TABLE_ROW_LIST, () => {
  //#region Mocked Responses
  let commandInfo: CommandInfo;
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validTableName = 'cr6c3_clitable';
  const validEntitySetName = 'cr6c3_clitables';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const tableResponse = {
    EntitySetName: 'cr6c3_clitables'
  };
  const rowsResponse = {
    "value": [
      {
        "cr6c3_clitablesid": "95c80273-3764-ed11-9561-000d3a4bbea4",
        "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f",
        "statecode": 0,
        "statuscode": 1,
        "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "modifiedon": "2022-11-14T16:14:45Z",
        "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "_modifiedby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "versionnumber": 1413873,
        "createdon": "2022-11-14T16:14:45Z",
        "cr6c3_columnone": "Test1",
        "overriddencreatedon": null,
        "importsequencenumber": null,
        "_modifiedonbehalfby_value": null,
        "utcconversiontimezonecode": null,
        "_createdonbehalfby_value": null,
        "_owningteam_value": null,
        "timezoneruleversionnumber": null
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
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DATAVERSE_TABLE_ROW_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if required options specified (entitySetName)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, entitySetName: validEntitySetName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, tableName: validTableName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves dataverse table rows with the entitySetName parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/${validEntitySetName}`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return rowsResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironment, entitySetName: validEntitySetName } });
    assert(loggerLogSpy.calledWith(rowsResponse.value));
  });

  it('retrieves dataverse table rows with the tableName parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions(LogicalName='${validTableName}')?$select=EntitySetName`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return tableResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/${validEntitySetName}`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return rowsResponse;
        }
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironment, tableName: validTableName } });
    assert(loggerLogSpy.calledWith(rowsResponse.value));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/${validEntitySetName}`)) {
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

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironment, entitySetName: validEntitySetName } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
