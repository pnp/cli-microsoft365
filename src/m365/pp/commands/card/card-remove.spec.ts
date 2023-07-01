import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
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
import ppCardGetCommand from './card-get.js';
import command from './card-remove.js';

describe(commands.CARD_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validName = 'CLI 365 Card';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
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
    assert.strictEqual(command.name, commands.CARD_REMOVE);
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

  it('prompts before removing the specified card owned by the currently signed-in user when confirm option not passed', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    await command.action(logger, {
      options: {
        environmentName: validEnvironment,
        id: validId
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified card owned by the currently signed-in user when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        environmentName: validEnvironment,
        id: validId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified card owned by the currently signed-in user when prompt confirmed', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === ppCardGetCommand) {
        return ({
          stdout: `{ "overwritetime": "1900-01-01T00:00:00Z", "_owningbusinessunit_value": "b419f090-fe22-ec11-b6e5-000d3ab596a1", "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238", "componentidunique": "e2b1d019-bd9a-491a-b888-693740711319", "_owninguser_value": "4f175d04-b952-ed11-bba2-000d3adf774e", "statecode": 0, "statuscode": 1, "ismanaged": false, "cardid": "${validId}", "_ownerid_value": "4f175d04-b952-ed11-bba2-000d3adf774e", "componentstate": 0, "modifiedon": "2022-10-29T08:22:46Z", "name": "${validName}", "_modifiedby_value": "4f175d04-b952-ed11-bba2-000d3adf774e", "versionnumber": 4463945, "createdon": "2022-10-29T08:22:46Z", "description": " ", "_createdby_value": "4f175d04-b952-ed11-bba2-000d3adf774e", "overriddencreatedon": null, "schemaversion": null, "importsequencenumber": null, "tags": null, "_modifiedonbehalfby_value": null, "utcconversiontimezonecode": null, "publishdate": null, "_createdonbehalfby_value": null, "hiddentags": null, "remixsourceid": null, "sizes": null, "coowners": null, "_owningteam_value": null, "publishsourceid": null, "timezoneruleversionnumber": null, "iscustomizable": { "Value": true, "CanBeChanged": true, "ManagedPropertyLogicalName": "iscustomizableanddeletable"}}`
        });
      }

      throw new CommandError('Unknown case');
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards(${validId})`) {
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
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: 'CLI 365 Card'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified card without confirmation prompt', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/cards(${validId})`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        id: validId,
        force: true
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
        environmentName: validEnvironment,
        id: validId,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
