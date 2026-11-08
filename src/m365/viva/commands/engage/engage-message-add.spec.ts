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
import command, { options } from './engage-message-add.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.ENGAGE_MESSAGE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const firstMessage: any = { messages: [{ "id": 470839661887488, "sender_id": 1496550646, "replied_to_id": null, "created_at": "2019/12/22 17:20:30 +0000", "network_id": 801445, "message_type": "update", "sender_type": "user", "url": "https://www.yammer.com/api/v1/messages/470839661887488", "web_url": "https://www.yammer.com/nubo.eu/messages/470839661887488", "group_id": 13114941440, "body": { "parsed": "send a letter to me", "plain": "send a letter to me", "rich": "send a letter to me" }, "thread_id": 470839661887488, "client_type": "O365 Api Auth", "client_url": "https://api.yammer.com", "system_message": false, "direct_message": false, "chat_client_sequence": null, "language": null, "notified_user_ids": [], "privacy": "public", "attachments": [], "liked_by": { "count": 0, "names": [] }, "content_excerpt": "send a letter to me", "group_created_id": 13114941440 }] };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.ENGAGE_MESSAGE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('repliedToId must be a number', () => {
    const actual = commandOptionsSchema.safeParse({ body: "test", repliedToId: 'nonumber' });
    assert.strictEqual(actual.success, false);
  });

  it('groupId must be a number', () => {
    const actual = commandOptionsSchema.safeParse({ body: "test", groupId: 'nonumber' });
    assert.strictEqual(actual.success, false);
  });

  it('networkId must be a number', () => {
    const actual = commandOptionsSchema.safeParse({ body: "test", networkId: 'nonumber' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if body and repliedToId set', () => {
    const actual = commandOptionsSchema.safeParse({ body: "test", repliedToId: 1234122 });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if body and directToUserIds set', () => {
    const actual = commandOptionsSchema.safeParse({ body: "test", directToUserIds: '1234125' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if body and groupId set', () => {
    const actual = commandOptionsSchema.safeParse({ body: "test", groupId: 1234123 });
    assert.strictEqual(actual.success, true);
  });

  it('fails if no groupId, directToUserId, or repliedToId is specified', () => {
    const actual = commandOptionsSchema.safeParse({ body: "test" });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      body: "test",
      groupId: 1234123,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('posts a message', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return firstMessage;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ body: "send a letter to me", groupId: 13114941440, debug: true }) });

    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 470839661887488);
  });

  it('posts a message handling json', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return firstMessage;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ body: "send a letter to me", groupId: 13114941440, debug: true, output: "json" }) });

    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 470839661887488);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ body: "test", groupId: 1234123 }) } as any), new CommandError('An error has occurred.'));
  });
});
