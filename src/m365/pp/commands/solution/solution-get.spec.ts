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
import { session } from '../../../../utils/session.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './solution-get.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.SOLUTION_GET, () => {
  let commandInfo: CommandInfo;
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = 'ee62fd63-e49e-4c09-80de-8fae1b9a427e';
  const validName = 'Default';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const solutionResponse = {
    "value": [
      {
        "solutionid": "00000001-0000-0000-0001-00000000009b",
        "uniquename": "Crc00f1",
        "version": "1.0.0.0",
        "installedon": "2021-10-01T21:54:14Z",
        "solutionpackageversion": null,
        "friendlyname": "Common Data Services Default Solution",
        "versionnumber": 860052,
        "publisherid": {
          "friendlyname": "CDS Default Publisher",
          "publisherid": "00000001-0000-0000-0000-00000000005a"
        }
      }
    ]
  };
  const solutionResponseText: any = {
    "uniquename": "Crc00f1",
    "version": "1.0.0.0",
    "publisher": "CDS Default Publisher"
  };

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
      powerPlatform.getSolutionByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SOLUTION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when no solution found', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').rejects(new Error(`The specified solution '${validName}' does not exist.`));

    await assert.rejects(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        name: validName
      }
    }), new CommandError(`The specified solution '${validName}' does not exist.`));
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

  it('retrieves a specific solution from power platform environment with the name parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').callsFake(async () => solutionResponse.value[0]);

    await command.action(logger, { options: { verbose: true, environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36', name: 'Default' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value[0]));
  });

  it('retrieves a specific solution from power platform environment with name parameter in format text', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').callsFake(async () => solutionResponse.value[0]);

    await command.action(logger, { options: { debug: true, environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36', name: 'Default', output: 'text' } });
    assert(loggerLogSpy.calledWith(solutionResponseText));
  });

  it('retrieves a specific solution from power platform environment with the id parameter', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions(ee62fd63-e49e-4c09-80de-8fae1b9a427e)?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          return solutionResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36', id: 'ee62fd63-e49e-4c09-80de-8fae1b9a427e' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value[0]));
  });

  it('retrieves a specific solution from power platform environment with id parameter in format text', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions(ee62fd63-e49e-4c09-80de-8fae1b9a427e)?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          return solutionResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36', id: 'ee62fd63-e49e-4c09-80de-8fae1b9a427e', output: 'text' } });
    assert(loggerLogSpy.calledWith(solutionResponseText));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: `Resource '' does not exist or one of its queried reference-property objects are not present`
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: '4be50206-9576-4237-8b17-38d8aadfaa36', name: 'Default' } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
