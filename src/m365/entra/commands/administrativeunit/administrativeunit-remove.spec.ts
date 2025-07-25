import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { telemetry } from '../../../../telemetry.js';
import request from '../../../../request.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import commands from '../../commands.js';
import command from './administrativeunit-remove.js';

describe(commands.ADMINISTRATIVEUNIT_REMOVE, () => {
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const displayName = 'European Division';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let promptIssued: boolean;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      entraAdministrativeUnit.getAdministrativeUnitByDisplayName,
      cli.handleMultipleResultsFound,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when neither id nor displayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when id is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: administrativeUnitId
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when displayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      displayName: displayName
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when id is not a valid UUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('removes the specified administrative unit by id without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: administrativeUnitId, force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes the specified administrative unit by displayName while prompting for confirmation', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').resolves({ id: administrativeUnitId, displayName: displayName });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { displayName: displayName } });
    assert(deleteRequestStub.called);
  });

  it('throws an error when administrative unit by id cannot be found', async () => {
    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `Resource '${administrativeUnitId}' does not exist or one of its queried reference-property objects are not present.`,
        innerError: {
          date: '2023-10-27T12:24:36',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: administrativeUnitId, force: true } }),
      new CommandError(error.error.message));
  });

  it('prompts before removing the specified administrative unit when confirm option not passed', async () => {
    await command.action(logger, { options: { id: administrativeUnitId } });

    assert(promptIssued);
  });

  it('aborts removing administrative unit when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: administrativeUnitId } });
    assert(deleteSpy.notCalled);
  });
});