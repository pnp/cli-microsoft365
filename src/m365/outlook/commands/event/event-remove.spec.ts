import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import { accessToken } from '../../../../utils/accessToken.js';
import command, { options } from './event-remove.js';
import { formatting } from '../../../../utils/formatting.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';

describe(commands.EVENT_REMOVE, () => {
  const eventId = 'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAENAAAiIsqMbYjsT5e-T7KzowPTAAAa_WKzAAA=';
  const userId = '6799fd1a-723b-4eb7-8e52-41ae530274ca';
  const userPrincipalName = 'john.doe@contoso.com';

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.post,
      accessToken.isAppOnlyAccessToken,
      accessToken.getUserIdFromAccessToken,
      accessToken.getUserNameFromAccessToken,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EVENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when userId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: eventId, userId: userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when userName is a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ id: eventId, userName: userPrincipalName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = commandOptionsSchema.safeParse({ id: eventId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: eventId, userId: 'invalid' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ id: eventId, userName: 'invalid' });
    assert.notStrictEqual(actual.success, true);
  });

  it('removes a specific event using delegated permissions without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events/${eventId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: eventId, force: true, verbose: true } });
    assert(deleteRequestStub.calledOnce);
  });

  it('permanently removes a specific event using delegated permissions without prompting for confirmation', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events/${eventId}/permanentDelete`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: eventId, permanent: true, force: true, verbose: true } });
    assert(postRequestStub.calledOnce);
  });

  it('removes a specific event using delegated permissions while prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events/${eventId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { id: eventId, verbose: true } });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes a specific event using delegated permissions from a calendar specified by userId matching the current user without prompting for confirmation', async () => {
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(userId);
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events/${eventId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: eventId, userId: userId, force: true, verbose: true } });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes a specific event using delegated permissions from a calendar specified by userName matching the current user without prompting for confirmation', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns(userPrincipalName);
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userPrincipalName)}/events/${eventId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: eventId, userName: userPrincipalName, force: true, verbose: true } });
    assert(deleteRequestStub.calledOnce);
  });

  it('throws an error when userId does not match current user when using delegated permissions', async () => {
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns('00000000-0000-0000-0000-000000000000');

    await assert.rejects(command.action(logger, { options: { id: eventId, userId: userId, force: true } }),
      new CommandError(`You can only remove your own events when using delegated permissions. The specified userId '${userId}' does not match the current user '00000000-0000-0000-0000-000000000000'.`));
  });

  it('throws an error when userName does not match current user when using delegated permissions', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns('other.user@contoso.com');

    await assert.rejects(command.action(logger, { options: { id: eventId, userName: userPrincipalName, force: true } }),
      new CommandError(`You can only remove your own events when using delegated permissions. The specified userName '${userPrincipalName}' does not match the current user 'other.user@contoso.com'.`));
  });

  it('succeeds when userName matches current user case-insensitively using delegated permissions', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns('John.Doe@Contoso.com');
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userPrincipalName)}/events/${eventId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: eventId, userName: userPrincipalName, force: true } });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes a specific event using application permissions from a calendar specified by userId without prompting for confirmation', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events/${eventId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: eventId, userId: userId, force: true, verbose: true } });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes a specific event using application permissions from a calendar specified by userName without prompting for confirmation', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userPrincipalName)}/events/${eventId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: eventId, userName: userPrincipalName, force: true, verbose: true } });
    assert(deleteRequestStub.calledOnce);
  });

  it('permanently removes a specific event using application permissions from a calendar specified by userId', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events/${eventId}/permanentDelete`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: eventId, userId: userId, permanent: true, force: true } });
    assert(postRequestStub.calledOnce);
  });

  it('throws an error when both userId and userName are not defined when removing an event using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { id: eventId } }),
      new CommandError(`The option 'userId' or 'userName' is required when removing an event using application permissions.`));
  });

  it('throws an error when both userId and userName are defined when removing an event using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { id: eventId, userId: userId, userName: userPrincipalName } }),
      new CommandError(`Both options 'userId' and 'userName' cannot be used together when removing an event using application permissions.`));
  });

  it('throws an error when both userId and userName are defined when removing an event using delegated permissions', async () => {
    await assert.rejects(command.action(logger, { options: { id: eventId, userId: userId, userName: userPrincipalName } }),
      new CommandError(`Both options 'userId' and 'userName' cannot be used together when removing an event using delegated permissions.`));
  });

  it('correctly handles API errors', async () => {
    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `The specified object was not found in the store., The process failed to get the correct properties.`,
        innerError: {
          date: '2023-10-27T12:24:36',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events/${eventId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: eventId, force: true } }),
      new CommandError(error.error.message));
  });

  it('prompts before removing the event when confirm option not passed', async () => {
    await command.action(logger, { options: { id: eventId } });

    assert(promptIssued);
  });

  it('aborts removing the event when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: eventId } });
    assert(deleteSpy.notCalled);
  });
});
