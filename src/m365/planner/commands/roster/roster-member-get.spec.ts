import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';
import { session } from '../../../../utils/session';
const command: Command = require('./roster-member-get');

describe(commands.ROSTER_MEMBER_GET, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validRosterId = 'iryDKm9VLku2HIoC2G-TX5gABJw0';
  const validUserId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validUserName = 'john.doe@contoso.com';
  const rosterMemberResponse = {
    "id": "c98ca8a9-1ae3-4709-ab65-5751f8d58694",
    "userId": "d242e467-bd06-4fa0-93c6-aea8aca9d90d",
    "tenantId": "8eca2a6b-80a4-4230-aca3-3781b92a179b",
    "roles": []
  };

  const userResponse = { value: [{ id: validUserId }] };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
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
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_MEMBER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        rosterId: validRosterId,
        userId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid upn', async () => {
    const actual = await command.validate({
      options: {
        rosterId: validRosterId,
        userName: 'John Doe'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (userId)', async () => {
    const actual = await command.validate({ options: { rosterId: validRosterId, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (userName)', async () => {
    const actual = await command.validate({ options: { rosterId: validRosterId, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
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
      options: {
        verbose: true,
        rosterId: validRosterId,
        userName: validUserName
      }
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
      options: {
        verbose: true,
        rosterId: validRosterId,
        userId: validUserId
      }
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
      options: {
        rosterId: validRosterId,
        userId: validUserId
      }
    }), new CommandError('The roster member cannot be found.'));
  });
});
