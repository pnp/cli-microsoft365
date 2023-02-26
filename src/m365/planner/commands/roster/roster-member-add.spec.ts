import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { CommandInfo } from '../../../../cli/CommandInfo';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { Cli } from '../../../../cli/Cli';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
const command: Command = require('./roster-member-add');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      request.post,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
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
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === AadUserGetCommand) {
        return ({
          stdout: `{ "businessPhones": [], "displayName": "John Doe", "givenName": null, "jobTitle": "CLI for Microsoft 365 contributor", "mail": "john.doe@contoso.com", "mobilePhone": null, "officeLocation": null, "preferredLanguage": null, "surname": "John", "userPrincipalName": "john.doe@contoso.com", "id": "${validUserId}" }`
        });
      }

      throw 'Unknown case';
    });

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
    sinon.stub(request, 'post').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: { rosterId: validRosterId, userId: validUserId }
    }), new CommandError('The requested item is not found.'));
  });
});