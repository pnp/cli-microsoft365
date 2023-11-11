import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './roster-member-remove.js';
import { ConfirmationConfig } from '../../../../utils/prompt.js';

describe(commands.ROSTER_MEMBER_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validRosterId = 'iryDKm9VLku2HIoC2G-TX5gABJw0';
  const validUserId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
  const validUserName = 'john.doe@contoso.com';
  const rosterMemberResponse = {
    value: [
      {
        id: '78ccf530-bbf0-47e4-aae6-da5f8c6fb142'
      },
      {
        id: 'eb77fbcf-6fe8-458b-985d-1747284793bc'
      }
    ]
  };

  const singleRosterMemberResponse = {
    value: [
      {
        id: '78ccf530-bbf0-47e4-aae6-da5f8c6fb142'
      }
    ]
  };
  const userResponse = { value: [{ id: validUserId }] };
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;

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
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.get,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_MEMBER_REMOVE);
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

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { rosterId: validRosterId, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { rosterId: validRosterId, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified roster member when force option not passed', async () => {
    await command.action(logger, {
      options: {
        rosterId: validRosterId,
        userId: validUserId
      }
    });

    assert(promptIssued);
  });

  it('prompts before removing the last roster member when force option not passed', async () => {
    let secondPromptConfirm: boolean = false;
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').callsFake(async (config: ConfirmationConfig) => {
      if (config.message === `Are you sure you want to remove member '${validUserId}'?`) {
        return true;
      }
      else {
        secondPromptConfirm = true;
        return false;
      }
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members?$select=Id`) {
        return singleRosterMemberResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        rosterId: validRosterId,
        userId: validUserId
      }
    });

    let promptIssued = false;

    if (secondPromptConfirm) {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified roster member when force option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');

    await command.action(logger, {
      options: {
        rosterId: validRosterId,
        userId: validUserId
      }
    });

    assert(deleteSpy.notCalled);
  });

  it('removes the last specified roster member when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'&$select=Id`) {
        return userResponse;
      }

      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members?$select=Id`) {
        return singleRosterMemberResponse;
      }

      throw 'Invalid request';
    });

    const deleteSpy = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members/${validUserId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        rosterId: validRosterId,
        userName: validUserName
      }
    });

    assert(deleteSpy.called);
  });

  it('removes the specified roster member when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'&$select=Id`) {
        return userResponse;
      }

      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members?$select=Id`) {
        return rosterMemberResponse;
      }

      throw 'Invalid request';
    });

    const deleteSpy = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members/${validUserId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        rosterId: validRosterId,
        userName: validUserName
      }
    });

    assert(deleteSpy.called);
  });

  it('removes the specified roster member without confirmation prompt', async () => {
    const deleteSpy = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members/${validUserId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        rosterId: validRosterId,
        userId: validUserId,
        force: true
      }
    });

    assert(deleteSpy.called);
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The roster member cannot be found.'
      }
    };
    sinon.stub(request, 'delete').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        rosterId: validRosterId,
        userId: validUserId,
        force: true
      }
    }), new CommandError('The roster member cannot be found.'));
  });
});
