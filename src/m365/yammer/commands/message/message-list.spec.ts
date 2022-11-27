import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./message-list');

describe(commands.MESSAGE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const firstMessageBatch: any = {
    messages: [
      { "sender_id": 1496550646, "replied_to_id": 1496550647, "id": 10123190123123, "thread_id": "", group_id: 11231123123, created_at: "2019/09/09 07:53:18 +0000", "content_excerpt": "message1", "body": { "plain": "message1 message is longer than 25 chars. Just for testing shortening" } },
      { "sender_id": 1496550640, "replied_to_id": "", "id": 10123190123124, "thread_id": "", group_id: "", created_at: "2019/09/08 07:53:18 +0000", "content_excerpt": "message2", "body": { "plain": "message2" } },
      { "sender_id": 1496550610, "replied_to_id": "", "id": 10123190123125, "thread_id": "", group_id: "", created_at: "2019/09/07 07:53:18 +0000", "content_excerpt": "message3", "body": { "plain": "message3" } },
      { "sender_id": 1496550630, "replied_to_id": "", "id": 10123190123126, "thread_id": "", group_id: 1123121, created_at: "2019/09/06 07:53:18 +0000", "content_excerpt": "message4", "body": { "plain": "message4" } },
      { "sender_id": 1496550646, "replied_to_id": "", "id": 10123190123127, "thread_id": "", group_id: 1123121, created_at: "2019/09/05 07:53:18 +0000", "content_excerpt": "message5", "body": { "plain": "message5" } }],
    meta: {
      older_available: true
    }
  };
  const secondMessageBatch: any = {
    messages: [
      { "sender_id": 1496550646, "replied_to_id": 1496550647, "id": 10123190123130, "thread_id": "", group_id: 11231123123, created_at: "2019/09/04 07:53:18 +0000", "content_excerpt": "message6", "body": { "plain": "message6" } },
      { "sender_id": 1496550640, "replied_to_id": "", "id": 10123190123131, "thread_id": "", group_id: "", created_at: "2019/09/03 07:53:18 +0000", "content_excerpt": "message7", "body": { "plain": "message7" } }],
    meta: {
      older_available: false
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.MESSAGE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'replied_to_id', 'thread_id', 'group_id', 'shortBody']);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('passes validation without parameters', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { limit: 10 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('limit must be a number', async () => {
    const actual = await command.validate({ options: { limit: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('olderThanId must be a number', async () => {
    const actual = await command.validate({ options: { olderThanId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('groupId must be a number', async () => {
    const actual = await command.validate({ options: { groupId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('threadId must be a number', async () => {
    const actual = await command.validate({ options: { threadId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('you are not allowed to use groupId and threadId at the same time', async () => {
    const actual = await command.validate({ options: { groupId: 123, threadId: 123 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('you cannot specify the feedType with groupId or threadId at the same time', async () => {
    const actual = await command.validate({ options: { feedType: 'All', threadId: 123 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('Fails in case FeedType is not correct', async () => {
    const actual = await command.validate({ options: { feedType: 'WrongValue' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('you are not allowed to use groupId and threadId and feedType at the same time', async () => {
    const actual = await command.validate({ options: { feedType: 'Private', groupId: 123, threadId: 123 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('returns messages without more results', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: {} } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from top feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/algo.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { feedType: 'Top' } } as any,);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from my feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/my_feed.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { feedType: 'My' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from following feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/following.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { feedType: 'Following' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from sent feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/sent.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { feedType: 'Sent' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from private feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/private.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { feedType: 'Private' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from received feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/received.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { feedType: 'Received' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from all feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { feedType: 'All' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from the group feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/in_group/123123.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { groupId: 123123 } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns messages from thread feed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/in_thread/123123.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { threadId: 123123 } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('returns all messages', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve(firstMessageBatch);
      }
      else {
        return Promise.resolve(secondMessageBatch);
      }
    });
    await command.action(logger, { options: { output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 7);
  });

  it('returns message with a specific limit', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve(firstMessageBatch);
      }
      else {
        return Promise.resolve(secondMessageBatch);
      }
    });
    await command.action(logger, { options: { limit: 6, output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 6);
  });

  it('handles error in loop', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve(firstMessageBatch);
      }
      else {
        return Promise.reject({
          "error": {
            "base": "An error has occurred."
          }
        });
      }
    });

    await assert.rejects(command.action(logger, { options: { output: 'json' } } as any), new CommandError('An error has occurred.'));
  });

  it('handles correct parameters older than', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json?older_than=10123190123128') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { olderThanId: 10123190123128, output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('handles correct parameters older than and threaded', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json?older_than=10123190123128&threaded=true') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { olderThanId: 10123190123128, threaded: true, output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });

  it('handles correct parameters with threaded', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json?threaded=true') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { threaded: true, output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 10123190123130);
  });
});