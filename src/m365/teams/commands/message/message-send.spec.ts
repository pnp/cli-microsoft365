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
import command from './message-send.js';

describe(commands.MESSAGE_SEND, () => {
  const teamId = '5f5d7b71-1161-44d8-bcc1-3da710eb4171';
  const channelId = '19:88f7e66a8dfe42be92db19505ae912a8@thread.skype';
  const message = 'Hello World';
  const messageSentResponse: any = {
    "@odata.context": `https://graph.microsoft.com/v1.0/$metadata#teams('${teamId}')/channels('19%3A88f7e66a8dfe42be92db19505ae912a8%40thread.tacv2')/messages/$entity`,
    "id": "1616990032035",
    "replyToId": null,
    "etag": "1616990032035",
    "messageType": "message",
    "createdDateTime": "2021-03-29T03:53:52.035Z",
    "lastModifiedDateTime": "2021-03-29T03:53:52.035Z",
    "lastEditedDateTime": null,
    "deletedDateTime": null,
    "subject": null,
    "summary": null,
    "chatId": null,
    "importance": "normal",
    "locale": "en-us",
    "webUrl": `https://teams.microsoft.com/l/message/19%3A88f7e66a8dfe42be92db19505ae912a8%40thread.tacv2/1616990032035?groupId=${teamId}&tenantId=2432b57b-0abd-43db-aa7b-16eadd115d34&createdTime=1616990032035&parentMessageId=1616990032035`,
    "policyViolation": null,
    "eventDetail": null,
    "from": {
      "application": null,
      "device": null,
      "user": {
        "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2",
        "displayName": "Robin Kline",
        "userIdentityType": "aadUser"
      }
    },
    "body": {
      "contentType": "html",
      "content": message
    },
    "channelIdentity": {
      "teamId": teamId,
      "channelId": channelId
    },
    "attachments": [],
    "mentions": [],
    "reactions": []
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_SEND);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        channelId: channelId,
        message: message
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the channelId is not valid', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        channelId: 'invalid',
        message: message
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        channelId: channelId,
        message: message
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('sends a message to a channel in a Microsoft Teams team', async () => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages`) {
        return messageSentResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        teamId: teamId,
        channelId: channelId,
        message: message
      }
    });

    assert(loggerLogSpy.calledWith(messageSentResponse));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      throw {
        error: {
          message: 'Channel does not belong to Team.'
        }
      };
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: teamId,
        channelId: channelId,
        message: message
      }
    }), new CommandError('Channel does not belong to Team.'));
  });
});