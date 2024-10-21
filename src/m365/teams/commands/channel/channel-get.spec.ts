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
import command from './channel-get.js';
import { teams } from '../../../../utils/teams.js';

describe(commands.CHANNEL_GET, () => {
  const teamId = '39958f28-eefb-4006-8f83-13b6ac2a4a7f';
  const teamName = 'Project Team';
  const channelId = '19:4eKaXAtxQJ4Xj3eUvCt4Zx5TPKBhF8jS7SfQYaA7lBY1@thread.tacv2';
  const channelName = 'channel1';
  const channelResponse = {
    id: channelId,
    displayName: channelName,
    description: null,
    email: "",
    webUrl: "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(teams, 'getTeamIdByDisplayName').resolves(teamId);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      teams.getTeamIdByDisplayName,
      teams.getChannelByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHANNEL_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        id: channelId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for a incorrect id missing leading 19:.', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly validates the when all options are valid', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        id: channelId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails to get channel information due to wrong channel id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}`) {
        throw {
          error: {
            code: 'ItemNotFound',
            message: 'Failed to execute Skype backend request GetThreadS2SRequest.',
            innerError: {
              'request-id': '4bebd0d2-d154-491b-b73f-d59ad39646fb',
              date: '2019-04-06T13:40:51'
            }
          }
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: teamId,
        id: channelId
      }
    } as any), new CommandError('Failed to execute Skype backend request GetThreadS2SRequest.'));
  });

  it('should get channel information for the Microsoft Teams team by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}`) {
        return channelResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: teamId,
        id: channelId
      }
    });
    assert(loggerLogSpy.calledWith(channelResponse));
  });

  it('should get primary channel information for the Microsoft Teams team by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/primaryChannel`) {
        return channelResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: teamId,
        primary: true
      }
    });
    assert(loggerLogSpy.calledWith(channelResponse));
  });

  it('should get channel information for the Microsoft Teams team by name', async () => {
    sinon.stub(teams, 'getChannelByDisplayName').withArgs(teamId, channelName).resolves(channelResponse);

    await command.action(logger, {
      options: {
        teamName: teamName,
        name: channelName
      }
    });
    assert(loggerLogSpy.calledWith(channelResponse));
  });

  it('should get primary channel information for the Microsoft Teams team by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/primaryChannel`) {
        return channelResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamName: teamName,
        primary: true
      }
    });
    assert(loggerLogSpy.calledWith(channelResponse));
  });

  it('should get channel information for the Microsoft Teams team', async () => {
    sinon.stub(teams, 'getChannelByDisplayName').withArgs(teamId, channelName).resolves(channelResponse);

    await command.action(logger, {
      options: {
        teamId: teamId,
        name: channelName
      }
    });
    assert(loggerLogSpy.calledWith(channelResponse));
  });
});
