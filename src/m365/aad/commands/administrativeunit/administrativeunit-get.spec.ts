import assert from 'assert';
import sinon from "sinon";
import auth from '../../../../Auth.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from "../../../../cli/Logger.js";
import commands from "../../commands.js";
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { Cli } from '../../../../cli/Cli.js';
import command from './administrativeunit-get.js';
import request from '../../../../request.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.ADMINISTRATIVEUNIT_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const administrativeUnitsReponse = {
    value: [
      {
        id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
        displayName: 'European Division',
        visibility: 'HiddenMembership'
      },
      {
        id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
        displayName: 'Asian Division',
        visibility: null
      }
    ]
  };
  const validId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const validDisplayName = 'European Division';
  const invalidDisplayName = 'European';

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified administrative unit by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${validId}`) {
        return administrativeUnitsReponse.value[0];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: validId } });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitsReponse.value[0]));
  });

  it('retrieves information about the specified administrative unit by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(validDisplayName)}'`) {
        return {
          value: [
            administrativeUnitsReponse.value[0]
          ]
        };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { displayName: validDisplayName } });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitsReponse.value[0]));
  });

  it('throws error message when no administrative unit was found by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, { options: { displayName: invalidDisplayName } }), new CommandError(`The specified administrative unit '${invalidDisplayName}' does not exist.`));
  });

  it('handles selecting single result when multiple administrative units with the specified displayName found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(validDisplayName)}'`) {
        return {
          value: [
            administrativeUnitsReponse.value[0],
            administrativeUnitsReponse.value[0]
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves({ id: validId, displayName: validDisplayName, visibility: 'HiddenMembership' });

    await command.action(logger, { options: { displayName: validDisplayName } });
    assert(loggerLogSpy.calledWith(administrativeUnitsReponse.value[0]));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { id: validId } }), new CommandError(errorMessage));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (displayName)', async () => {
    const actual = await command.validate({ options: { displayName: validDisplayName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});