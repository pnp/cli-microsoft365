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
import command from './message-restore.js';
import { settingsNames } from '../../../../settingsNames.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { teams } from '../../../../utils/teams.js';

describe(commands.MESSAGE_RESTORE, () => {
  const messageId = '1540911392778';
  const teamId = '5f5d7b71-1161-44d8-bcc1-3da710eb4171';
  const channelId = '19:00000000000000000000000000000000@thread.skype';
  const teamName = 'Team Name';
  const channelName = 'Channel Name';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      return settingName === settingsNames.prompt ? false : defaultValue;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_RESTORE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if teamId or teamName options are not passed', async () => {
    const actual = await command.validate({
      options: {
        id: messageId,
        channelId: channelId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if teamId and teamName options are both passed', async () => {
    const actual = await command.validate({
      options: {
        id: messageId,
        teamId: teamId,
        teamName: teamName,
        channelId: channelId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelId or channelName options are not passed', async () => {
    const actual = await command.validate({
      options: {
        id: messageId,
        teamId: teamId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelId and channelName options are both passed', async () => {
    const actual = await command.validate({
      options: {
        id: messageId,
        teamId: teamId,
        channelName: channelName,
        channelId: channelId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: "5f5d7b71-1161-44",
        channelId: channelId,
        id: messageId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        channelId: channelId,
        id: messageId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation for an incorrect channelId missing leading 19:.', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        id: messageId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for an incorrect channelId missing trailing @thread.skype.', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread',
        id: messageId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('restores the specified message', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages/${messageId}/undoSoftDelete`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: teamId,
        channelId: channelId,
        id: messageId
      }
    });
  });

  it('restores the specified message by team name and channel name (debug)', async () => {
    sinon.stub(teams, 'getChannelIdByDisplayName').resolves(channelId);
    sinon.stub(teams, 'getTeamIdByDisplayName').resolves(teamId);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages/${messageId}/undoSoftDelete`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        teamName: teamName,
        channelName: channelName,
        id: messageId
      }
    });
  });

  it('correctly handles error when retrieving a message', async () => {
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

    await assert.rejects(command.action(logger, {
      options: {
        teamId: teamId,
        channelId: channelId,
        id: messageId
      }
    }), new CommandError('An error has occurred'));
  });
});