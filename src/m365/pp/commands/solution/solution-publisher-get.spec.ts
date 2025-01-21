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
import command from './solution-publisher-get.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.SOLUTION_PUBLISHER_GET, () => {
  let commandInfo: CommandInfo;
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = 'd21aab70-79e7-11dd-8874-00188b01e34f';
  const validName = 'MicrosoftCorporation';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const publisherResponse = {
    "value": [
      {
        "publisherid": "d21aab70-79e7-11dd-8874-00188b01e34f",
        "uniquename": "MicrosoftCorporation",
        "friendlyname": "MicrosoftCorporation",
        "versionnumber": 1226559,
        "isreadonly": false,
        "customizationprefix": "",
        "customizationoptionvalueprefix": 0
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
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
    assert.strictEqual(command.name, commands.SOLUTION_PUBLISHER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when no publisher found', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$filter=friendlyname eq '${validName}'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
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
    }), new CommandError(`The specified publisher '${validName}' does not exist.`));
  });

  it('fails validation if the id is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, name: validName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves a specific publisher from power platform environment with the name parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$filter=friendlyname eq '${validName}'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          return publisherResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: validEnvironment, name: validName } });
    assert(loggerLogSpy.calledWith(publisherResponse.value[0]));
  });

  it('retrieves a specific publisher from power platform environment with the id parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers(${validId})?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          return publisherResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: validEnvironment, id: validId } });
    assert(loggerLogSpy.calledWith(publisherResponse.value[0]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$filter=friendlyname eq '${validName}'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
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
