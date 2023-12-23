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
import command from './card-clone.js';
import ppCardGetCommand from './card-get.js';

describe(commands.CARD_CLONE, () => {
  let commandInfo: CommandInfo;
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validName = 'CLI 365 Card';
  const validNewName = 'new CLI 365 Card';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const cardResponse = {
    "CardIdClone": "80cff342-ddf1-4633-aec1-6d3d131b29e0"
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
      request.post,
      powerPlatform.getDynamicsInstanceApiUrl,
      cli.promptForConfirmation,
      cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CARD_CLONE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: 'Invalid GUID',
        newName: validName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId, newName: validNewName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, name: validName, newName: validNewName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('clones the specified card owned by the currently signed-in user based on the name', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === ppCardGetCommand) {
        return ({
          stdout: `{ "overwritetime": "1900-01-01T00:00:00Z", "_owningbusinessunit_value": "b419f090-fe22-ec11-b6e5-000d3ab596a1", "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238", "componentidunique": "e2b1d019-bd9a-491a-b888-693740711319", "_owninguser_value": "4f175d04-b952-ed11-bba2-000d3adf774e", "statecode": 0, "statuscode": 1, "ismanaged": false, "cardid": "${validId}", "_ownerid_value": "4f175d04-b952-ed11-bba2-000d3adf774e", "componentstate": 0, "modifiedon": "2022-10-29T08:22:46Z", "name": "${validName}", "_modifiedby_value": "4f175d04-b952-ed11-bba2-000d3adf774e", "versionnumber": 4463945, "createdon": "2022-10-29T08:22:46Z", "description": " ", "_createdby_value": "4f175d04-b952-ed11-bba2-000d3adf774e", "overriddencreatedon": null, "schemaversion": null, "importsequencenumber": null, "tags": null, "_modifiedonbehalfby_value": null, "utcconversiontimezonecode": null, "publishdate": null, "_createdonbehalfby_value": null, "hiddentags": null, "remixsourceid": null, "sizes": null, "coowners": null, "_owningteam_value": null, "publishsourceid": null, "timezoneruleversionnumber": null, "iscustomizable": { "Value": true, "CanBeChanged": true, "ManagedPropertyLogicalName": "iscustomizableanddeletable"}}`
        });
      }

      throw new CommandError('Unknown case');
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/CardCreateClone` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "CardId": validId,
          "CardName": validNewName
        })) {
        return cardResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: validName,
        newName: validNewName
      }
    });
    assert(loggerLogSpy.calledWith(cardResponse));
  });

  it('clones the specified card owned by the currently signed-in user based on the id', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/CardCreateClone` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "CardId": validId,
          "CardName": validNewName
        })) {
        return cardResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        id: validId,
        newName: validNewName
      }
    });
    assert(loggerLogSpy.calledWith(cardResponse));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = `The environment '${validEnvironment}' could not be retrieved. See the inner exception for more details: undefined`;
    sinon.stub(request, 'get').callsFake(async () => { throw errorMessage; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        name: validName,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
