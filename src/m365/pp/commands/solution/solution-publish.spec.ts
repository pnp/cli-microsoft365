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
import command from './solution-publish.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.SOLUTION_PUBLISH, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '00000001-0000-0000-0001-00000000009b';
  const validName = 'Solution name';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const validParameterXmlData = {
    ParameterXml: '<importexportxml><canvasapps><canvasapp>new_test_eb178</canvasapp></canvasapps><entities><entity>new_test</entity></entities></importexportxml>'
  };
  const validSolutionComponentsResult = {
    value: [
      {
        'msdyn_componentlogicalname': 'canvasapp',
        'msdyn_name': 'new_test_eb178'
      },
      {
        'msdyn_componentlogicalname': 'entity',
        'msdyn_name': 'new_test'
      }
    ]
  };
  const solutionResponse = {
    solutionid: validId,
    uniquename: validName,
    version: '1.0.0.0',
    installedon: '2021-10-01T21:54:14Z',
    solutionpackageversion: null,
    friendlyname: validName,
    versionnumber: 860052,
    publisherid: {
      friendlyname: 'CDS Default Publisher',
      publisherid: '00000001-0000-0000-0000-00000000005a'
    }
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl,
      cli.promptForConfirmation,
      cli.executeCommandWithOutput,
      powerPlatform.getSolutionByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SOLUTION_PUBLISH);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [{ options: ['id', 'name'] }]);
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

  it('publishes the components of a specified solution owned by the currently signed-in user', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${validId})&$select=msdyn_componentlogicalname,msdyn_name&$orderby=msdyn_componentlogicalname asc&api-version=9.1`) {
        return validSolutionComponentsResult;
      }

      throw `Invalid request with opts ${JSON.stringify(opts)}`;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/PublishXml`) {
        if (JSON.stringify(opts.data) === JSON.stringify(validParameterXmlData)) {
          return;
        }
      }

      throw `Invalid request with opts ${JSON.stringify(opts)}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        name: validName
      }
    }));
  });

  it('publishes the components of a specified solution owned by the currently signed-in user and waits for completion', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${validId})&$select=msdyn_componentlogicalname,msdyn_name&$orderby=msdyn_componentlogicalname asc&api-version=9.1`) {
        return validSolutionComponentsResult;
      }

      throw `Invalid request with opts ${JSON.stringify(opts)}`;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/PublishXml`) {
        if (JSON.stringify(opts.data) === JSON.stringify(validParameterXmlData)) {
          return;
        }
      }

      throw `Invalid request with opts ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        name: validName,
        wait: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        id: validId
      }
    }), new CommandError(errorMessage));
  });
});