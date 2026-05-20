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
import command, { options } from './engage-role-member-add.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { entraUser } from '../../../../utils/entraUser.js';

describe(commands.ENGAGE_ROLE_MEMBER_ADD, () => {
  const roleId = 'ec759127-089f-4f91-8dfc-03a30b51cb38';
  const roleName = 'Network Admin';
  const userId = 'a1b2c3d4-e5f6-4789-9012-3456789abcde';
  const userName = 'john.doe@contoso.com';
  const addRoleMemberResponse = {
    "@odata.type": "#microsoft.graph.engagementRoleMember",
    "id": "a40473a5-0fb4-a250-e029-f6fe33d07733",
    "userId": userId,
    "createdDateTime": "2026-04-15T14:03:00Z"
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      request.post,
      entraUser.getUserIdByUpn,
      vivaEngage.getRoleIdByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_ROLE_MEMBER_ADD);
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

  it('adds the user specified by id to Viva Engage role specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles/${roleId}/members`) {
        return addRoleMemberResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, roleId: roleId, userId: userId }) });
    assert(loggerLogSpy.calledWith(addRoleMemberResponse));
  });

  it('adds the user specified by name to Viva Engage role specified by name', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    sinon.stub(vivaEngage, 'getRoleIdByName').withArgs(roleName).resolves(roleId);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles/${roleId}/members`) {
        return addRoleMemberResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, roleName: roleName, userName: userName }) });
    assert(loggerLogSpy.calledWith(addRoleMemberResponse));
  });

  it('handles error when adding a user to a Viva Engage role failed', async () => {
    sinon.stub(request, 'post').rejects({ error: { message: 'An error has occurred' } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ roleId: roleId, userId: userId }) }),
      new CommandError('An error has occurred')
    );
  });
});