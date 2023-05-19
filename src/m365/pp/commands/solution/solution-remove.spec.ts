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
import * as PpSolutionGetCommand from './solution-get';
const command: Command = require('./solution-remove');

describe(commands.SOLUTION_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '00000001-0000-0000-0001-00000000009b';
  const validName = 'Solution name';
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
    assert.strictEqual(command.name.startsWith(commands.SOLUTION_REMOVE), true);
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

  it('prompts before removing the specified solution owned by the currently signed-in user when confirm option not passed', async () => {
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

  it('aborts removing the specified solution owned by the currently signed-in user when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');

    await command.action(logger, {
      options: {
        environment: validEnvironment,
        id: validId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified solution owned by the currently signed-in user when prompt confirmed', async () => {
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

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/solutions(${validId})`) {
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

  it('removes the specified solution owned by the currently signed-in user without prompt for confirm', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/solutions(${validId})`) {
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
