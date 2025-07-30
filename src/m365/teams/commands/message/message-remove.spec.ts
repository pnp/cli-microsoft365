import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './message-remove.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { teams } from '../../../../utils/teams.js';

describe(commands.MESSAGE_REMOVE, () => {
  const channelId = '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype';
  const channelName = 'channelName';
  const teamId = 'd66b8110-fcad-49e8-8159-0d488ddb7656';
  const teamName = 'Team Name';
  const messageId = '157836366';

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').resolves();
    sinon.stub(teams, 'getTeamIdByDisplayName').resolves(teamId);
    sinon.stub(teams, 'getChannelIdByDisplayName').resolves(channelId);
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({ options: { teamId: 'invalid', channelId: channelId, id: messageId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the channel id is not a valid channel id', async () => {
    const actual = await command.validate({ options: { teamId: teamId, channelId: 'invalid', id: messageId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the channel id and team id both are valid', async () => {
    const actual = await command.validate({ options: { teamId: teamId, channelId: channelId, id: messageId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the channel name and team name are set', async () => {
    const actual = await command.validate({ options: { teamName: teamName, channelName: channelName, id: messageId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified message when force option not passed', async () => {
    await command.action(logger, { options: { id: messageId, teamId: teamId, channelId: channelId } });
    assert(promptIssued);
  });

  it('aborts removing the specified message when force option not passed and prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();
    await command.action(logger, { options: { id: messageId, teamId: teamId, channelId: channelId } });
    assert(postStub.notCalled);
  });

  it('removes the specified message when teamId, channelId and force option passed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${formatting.encodeQueryParameter(channelId)}/messages/${messageId}/softDelete`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: messageId, teamId: teamId, channelId: channelId, force: true, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('removes the specified message when teamName, channelName is used and prompt is confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${formatting.encodeQueryParameter(channelId)}/messages/${messageId}/softDelete`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { id: messageId, teamName: teamName, channelName: channelName, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('throws error when the message we are trying to delete is not found', async () => {
    const error = {
      error: {
        error: {
          code: 'NotFound',
          message: 'NotFound',
          innerError: {
            code: '1',
            message: `MessageNotFound-Message does not exist in the thread: ColdStoreNotSupportedForMessageException:ColdStoreMessageOperations is not supported if cutOffColdStoreEpoch is not provided. (msgVersion:${messageId}, cutoff:1707944123071)`,
            date: '2024-02-21T20:55:23',
            'request-id': 'fe227b45-0b96-47c2-bac4-3d5e17dfc70d',
            'client-request-id': 'fe227b45-0b96-47c2-bac4-3d5e17dfc70d'
          }
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: { id: messageId, teamName: teamName, channelName: channelName, force: true, verbose: true } }),
      new CommandError('The message was not found in the Teams channel.'));
  });

  it('correctly handles generic error when removing message', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };

    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: { id: messageId, channelId: channelId, teamName: teamName, force: true, verbose: true } } as any),
      new CommandError('An error has occurred'));
  });
});
