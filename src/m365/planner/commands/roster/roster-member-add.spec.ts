import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './roster-member-add.js';

describe(commands.ROSTER_MEMBER_ADD, () => {
  let commandInfo: CommandInfo;
  const rosterMemberResponse = {
    "id": "b3a1be03-54a5-43d2-b4fb-6562fe9bec0b",
    "userId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
    "tenantId": "5b7b813c-2339-48cd-8c51-bd4fcb269420",
    "roles": []
  };
  const validRosterId = "iryDKm9VLku2HIoC2G-TX5gABJw0";
  const validUserId = "2056d2f6-3257-4253-8cfc-b73393e414e5";
  const validUserName = "john.doe@contoso.com";

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
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
      request.post,
      aadUser.getUserIdByUpn
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_MEMBER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        rosterId: validRosterId,
        userId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userName is not a valid upn', async () => {
    const actual = await command.validate({
      options: {
        rosterId: validRosterId,
        userName: 'Invalid upn'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { rosterId: validRosterId, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { rosterId: validRosterId, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly adds a new roster member by userId', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return rosterMemberResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { rosterId: validRosterId, userId: validUserId } });
    assert(loggerLogSpy.calledWith(rosterMemberResponse));
  });

  it('adds a new member to the roster by userName', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').resolves(validUserId);

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return rosterMemberResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { verbose: true, rosterId: validRosterId, userName: validUserName } });
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
      options: { rosterId: validRosterId, userId: validUserId }
    }), new CommandError('The requested item is not found.'));
  });
});