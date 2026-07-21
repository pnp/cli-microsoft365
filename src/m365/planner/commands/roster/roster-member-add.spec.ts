import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './roster-member-add.js';

describe(commands.ROSTER_MEMBER_ADD, () => {
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const rosterMemberResponse = {
    id: 'b3a1be03-54a5-43d2-b4fb-6562fe9bec0b',
    userId: '2056d2f6-3257-4253-8cfc-b73393e414e5',
    tenantId: '5b7b813c-2339-48cd-8c51-bd4fcb269420',
    roles: []
  };
  const validRosterId = 'iryDKm9VLku2HIoC2G-TX5gABJw0';
  const validUserId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validUserName = 'john.doe@contoso.com';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
      entraUser.getUserIdByUpn
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_MEMBER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid guid.', () => {
    const actual = commandOptionsSchema.safeParse({
      rosterId: validRosterId,
      userId: 'Invalid GUID'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when userName is not a valid upn', () => {
    const actual = commandOptionsSchema.safeParse({
      rosterId: validRosterId,
      userName: 'Invalid upn'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when neither userId nor userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      rosterId: validRosterId
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      rosterId: validRosterId,
      userId: validUserId,
      userName: validUserName
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if required options specified (id)', () => {
    const actual = commandOptionsSchema.safeParse({ rosterId: validRosterId, userId: validUserId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if required options specified (name)', () => {
    const actual = commandOptionsSchema.safeParse({ rosterId: validRosterId, userName: validUserName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      rosterId: validRosterId,
      userId: validUserId,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('correctly adds a new roster member by userId', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return rosterMemberResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ rosterId: validRosterId, userId: validUserId }) });
    assert(loggerLogSpy.calledWith(rosterMemberResponse));
  });

  it('adds a new member to the roster by userName', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').resolves(validUserId);

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return rosterMemberResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, rosterId: validRosterId, userName: validUserName }) });
    assert(loggerLogSpy.calledWith(rosterMemberResponse));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The requested item is not found.'
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({ rosterId: validRosterId, userId: validUserId })
    }), new CommandError('The requested item is not found.'));
  });
});
