import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './roledefinition-remove.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { z } from 'zod';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';

describe(commands.ROLEDEFINITION_REMOVE, () => {
  const roleId = 'abcd1234-de71-4623-b4af-96380a352509';
  const roleDisplayName = 'Bitlocker Keys Reader';

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
      request.get,
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROLEDEFINITION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both id and displayName are provided', () => {
    const actual = commandOptionsSchema.safeParse({
      id: roleId,
      displayName: roleDisplayName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither id nor displayName is provided', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('prompts before removing the role definition when confirm option not passed', async () => {
    await command.action(logger, { options: { id: roleId } });

    assert(promptIssued);
  });

  it('aborts removing the role definition when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: roleId } });
    assert(deleteSpy.notCalled);
  });

  it('removes the role definition specified by id without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: roleId, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes the role definition specified by displayName while prompting for confirmation', async () => {
    sinon.stub(roleDefinition, 'getRoleDefinitionByDisplayName').resolves({ id: roleId });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { displayName: roleDisplayName } });
    assert(deleteRequestStub.called);
  });

  it('handles error when role definition specified by id was not found', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${roleId}`) {
        throw {
          error:
          {
            code: 'Request_ResourceNotFound',
            message: `Resource '${roleId}' does not exist or one of its queried reference-property objects are not present.`
          }
        };
      }
      throw `Invalid request`;
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(
      command.action(logger, { options: { id: roleId } }),
      new CommandError(`Resource '${roleId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });
});