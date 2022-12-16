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
const command: Command = require('./message-add');

describe(commands.MESSAGE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const firstMessage: any = { messages: [{ "id": 470839661887488, "sender_id": 1496550646, "replied_to_id": null, "created_at": "2019/12/22 17:20:30 +0000", "network_id": 801445, "message_type": "update", "sender_type": "user", "url": "https://www.yammer.com/api/v1/messages/470839661887488", "web_url": "https://www.yammer.com/nubo.eu/messages/470839661887488", "group_id": 13114941440, "body": { "parsed": "send a letter to me", "plain": "send a letter to me", "rich": "send a letter to me" }, "thread_id": 470839661887488, "client_type": "O365 Api Auth", "client_url": "https://api.yammer.com", "system_message": false, "direct_message": false, "chat_client_sequence": null, "language": null, "notified_user_ids": [], "privacy": "public", "attachments": [], "liked_by": { "count": 0, "names": [] }, "content_excerpt": "send a letter to me", "group_created_id": 13114941440 }] };

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
    assert.strictEqual(command.name.startsWith(commands.MESSAGE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id']);
  });

  it('repliedToId must be a number', async () => {
    const actual = await command.validate({ options: { body: "test", repliedToId: 'nonumber' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('groupId must be a number', async () => {
    const actual = await command.validate({ options: { body: "test", groupId: 'nonumber' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('networkId must be a number', async () => {
    const actual = await command.validate({ options: { body: "test", networkId: 'nonumber' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if body and repliedToId set', async () => {
    const actual = await command.validate({ options: { body: "test", repliedToId: 1234122 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if body and directToUserIds set', async () => {
    const actual = await command.validate({ options: { body: "test", directToUserIds: 1234125 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if body and groupId set', async () => {
    const actual = await command.validate({ options: { body: "test", groupId: 1234123 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails if no groupId, directToUserId, or repliedToId is specified', async () => {
    const actual = await command.validate({ options: { body: "test" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('posts a message', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(firstMessage);
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { body: "send a letter to me", groupId: 13114941440, debug: true } } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 470839661887488);
  });

  it('posts a message handling json', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(firstMessage);
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { body: "send a letter to me", groupId: 13114941440, debug: true, output: "json" } } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 470839661887488);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });
});
