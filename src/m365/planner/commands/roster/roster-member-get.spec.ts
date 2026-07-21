import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './roster-member-get.js';

describe(commands.ROSTER_MEMBER_GET, () => {
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  //#region Mocked Responses
  const validRosterId = 'iryDKm9VLku2HIoC2G-TX5gABJw0';
  const validUserId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validUserName = 'john.doe@contoso.com';
  const rosterMemberResponse = {
    id: 'c98ca8a9-1ae3-4709-ab65-5751f8d58694',
    userId: 'd242e467-bd06-4fa0-93c6-aea8aca9d90d',
    tenantId: '8eca2a6b-80a4-4230-aca3-3781b92a179b',
    roles: []
  };

  const userResponse = { value: [{ id: validUserId }] };
  //#endregion

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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_MEMBER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid guid', () => {
    const actual = commandOptionsSchema.safeParse({
      rosterId: validRosterId,
      userId: 'Invalid GUID'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userName is not a valid upn', () => {
    const actual = commandOptionsSchema.safeParse({
      rosterId: validRosterId,
      userName: 'John Doe'
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

  it('passes validation if required options specified (userId)', () => {
    const actual = commandOptionsSchema.safeParse({ rosterId: validRosterId, userId: validUserId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if required options specified (userName)', () => {
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

  it('gets the specified roster member by userName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'&$select=Id`) {
        return userResponse;
      }

      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members/${validUserId}`) {
        return rosterMemberResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        rosterId: validRosterId,
        userName: validUserName
      })
    });

    assert(loggerLogSpy.calledWith(rosterMemberResponse));
  });

  it('gets the specified roster member by userId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members/${validUserId}`) {
        return rosterMemberResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        rosterId: validRosterId,
        userId: validUserId
      })
    });

    assert(loggerLogSpy.calledWith(rosterMemberResponse));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The roster member cannot be found.'
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        rosterId: validRosterId,
        userId: validUserId
      })
    }), new CommandError('The roster member cannot be found.'));
  });
});
