import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as PpSolutionGetCommand from './solution-get';
const command: Command = require('./solution-publish');

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
  //#endregion

  let log: string[];
  let logger: Logger;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
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
    assert.strictEqual(command.name.startsWith(commands.SOLUTION_PUBLISH), true);
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

  it('publishes the components of a specified solution owned by the currently signed-in user', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === PpSolutionGetCommand) {
        return ({
          stdout: `{
            "solutionid": "${validId}",
            "uniquename": "${validName}",
            "version": "1.0.0.0",
            "installedon": "2022-10-30T13:59:26Z",
            "solutionpackageversion": null,
            "friendlyname": "${validName}",
            "versionnumber": 1209676,
            "publisherid": {
              "friendlyname": "Default Publisher for org1547b730",
              "publisherid": "d21aab71-79e7-11dd-8874-00188b01e34f"
            }
          }`
        });
      }

      throw new CommandError('Unknown case');
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${validId})&$select=msdyn_componentlogicalname,msdyn_name&$orderby=msdyn_componentlogicalname asc&api-version=9.1`) {
        return validSolutionComponentsResult;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/PublishXml`) {
        if (JSON.stringify(opts.data) === JSON.stringify(validParameterXmlData)) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        environment: validEnvironment,
        name: validName
      }
    }));
  });

  it('publishes the components of a specified solution owned by the currently signed-in user and waits for completion', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === PpSolutionGetCommand) {
        return ({
          stdout: `{
            "solutionid": "${validId}",
            "uniquename": "${validName}",
            "version": "1.0.0.0",
            "installedon": "2022-10-30T13:59:26Z",
            "solutionpackageversion": null,
            "friendlyname": "${validName}",
            "versionnumber": 1209676,
            "publisherid": {
              "friendlyname": "Default Publisher for org1547b730",
              "publisherid": "d21aab71-79e7-11dd-8874-00188b01e34f"
            }
          }`
        });
      }

      throw new CommandError('Unknown case');
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${validId})&$select=msdyn_componentlogicalname,msdyn_name&$orderby=msdyn_componentlogicalname asc&api-version=9.1`) {
        return validSolutionComponentsResult;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/PublishXml`) {
        if (JSON.stringify(opts.data) === JSON.stringify(validParameterXmlData)) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environment: validEnvironment,
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
        environment: validEnvironment,
        id: validId
      }
    }), new CommandError(errorMessage));
  });
});