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
const command: Command = require('./message-send');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
      request.post
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