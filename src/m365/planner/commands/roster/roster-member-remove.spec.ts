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
import { powerPlatform } from '../../../../utils/powerPlatform';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./roster-member-remove');

describe(commands.ROSTER_MEMBER_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validRosterId = "iryDKm9VLku2HIoC2G-TX5gABJw0";
  const validUserId = "2056d2f6-3257-4253-8cfc-b73393e414e5";
  const validUserName = "john.doe@contoso.com";
  const rosterMemberResponse = {
    value: [
      {
        id: "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
        userId: "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
        tenantId: "0cac6cda-2e04-4a3d-9c16-9c91470d7022",
        roles: []
      },
      {
        id: "eb77fbcf-6fe8-458b-985d-1747284793bc",
        userId: "eb77fbcf-6fe8-458b-985d-1747284793bc",
        tenantId: "0cac6cda-2e04-4a3d-9c16-9c91470d7022",
        roles: []
      }
    ]
  };
  const userResponse = { value: [{ "id": validUserId, "businessPhones": ["+1 425 555 0100"], "displayName": "Aarif Sherzai", "givenName": "Aarif", "jobTitle": "Administrative", "mail": null, "mobilePhone": "+1 425 555 0100", "officeLocation": null, "preferredLanguage": null, "surname": "Sherzai", "userPrincipalName": validUserName }] };
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_MEMBER_REMOVE);
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

  it('prompts before removing the specified roster member when confirm option not passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return rosterMemberResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        rosterId: validRosterId,
        userid: validUserId
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing the last roster member when confirm option not passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return ({ value: [rosterMemberResponse.value[0]] });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        rosterId: validRosterId,
        userid: validUserId
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified specified roster member when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return rosterMemberResponse;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        rosterId: validRosterId,
        userid: validUserId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified roster member when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'`) {
        return userResponse;
      }

      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return rosterMemberResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members/${validUserId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, {
      options: {
        debug: true,
        rosterId: validRosterId,
        userName: validUserName
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified roster member without confirmation prompt', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members/${validUserId}`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`) {
        return rosterMemberResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        rosterId: validRosterId,
        userId: validUserId,
        confirm: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('fails to get user for roster when user with provided user name does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'`) > -1) {
        return ({ value: [] });
      }

      throw `The specified user with user name ${validUserName} does not exist`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        rosterId: validRosterId,
        userName: validUserName,
        confirm: true
      }
    }), new CommandError(`The specified user with user name ${validUserName} does not exist`));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The roster member cannot be found.'
      }
    };
    sinon.stub(request, 'delete').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: {
        rosterId: validRosterId,
        userId: validUserId,
        confirm: true
      }
    }), new CommandError('The roster member cannot be found.'));
  });
});
