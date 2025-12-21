import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './engage-role-member-remove.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { entraUser } from '../../../../utils/entraUser.js';

describe(commands.ENGAGE_ROLE_MEMBER_REMOVE, () => {
  const roleId = 'ec759127-089f-4f91-8dfc-03a30b51cb38';
  const roleName = 'Network Admin';
  const userId = 'a1b2c3d4-e5f6-4789-9012-3456789abcde';
  const userName = 'john.doe@contoso.com';

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
      entraUser.getUserIdByUpn,
      vivaEngage.getRoleIdByName,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_ROLE_MEMBER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if roleId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: 'invalid',
      userId: userId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if roleId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId,
      userId: userId
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if roleName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleName: roleName,
      userId: userId
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both roleId and roleName are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId,
      roleName: roleName,
      userId: userId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither roleId nor roleName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId,
      userId: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if userId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId,
      userId: userId
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId,
      userName: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId,
      userName: userName
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId,
      userId: userId,
      userName: userName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither userId nor userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      roleId: roleId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('removes the user specified by id from Viva Engage role specified by id without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/roles/${roleId}/members/${userId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ roleId: roleId, userId: userId, force: true }) });
    assert(deleteRequestStub.called);
  });

  it('removes the user specified by name from Viva Engage role specified by id while prompting for confirmation', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/roles/${roleId}/members/${userId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ roleId: roleId, userName: userName }) });
    assert(deleteRequestStub.called);
  });

  it('removes the user specified by id from Viva Engage role specified by name while prompting for confirmation', async () => {
    sinon.stub(vivaEngage, 'getRoleIdByName').withArgs(roleName).resolves(roleId);

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/roles/${roleId}/members/${userId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ roleName: roleName, userId: userId }) });
    assert(deleteRequestStub.called);
  });

  it('removes the user specified by name from Viva Engage role specified by name while prompting for confirmation', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(vivaEngage, 'getRoleIdByName').withArgs(roleName).resolves(roleId);

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/roles/${roleId}/members/${userId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ roleName: roleName, userName: userName, verbose: true }) });
    assert(deleteRequestStub.called);
  });

  it('handles error when removing a user from a Viva Engage role failed', async () => {
    sinon.stub(request, 'delete').rejects({ error: { message: 'An error has occurred' } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ roleId: roleId, userId: userId, force: true }) }),
      new CommandError('An error has occurred')
    );
  });

  it('prompts before removing a member from a Viva Engage role when confirm option not passed', async () => {
    await command.action(logger, { options: { roleId: roleId, userId: userId } });

    assert(promptIssued);
  });
});