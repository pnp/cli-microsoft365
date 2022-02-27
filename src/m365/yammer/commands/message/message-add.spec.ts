import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./message-add');

describe(commands.MESSAGE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const firstMessage: any = { messages: [{ "id": 470839661887488, "sender_id": 1496550646, "replied_to_id": null, "created_at": "2019/12/22 17:20:30 +0000", "network_id": 801445, "message_type": "update", "sender_type": "user", "url": "https://www.yammer.com/api/v1/messages/470839661887488", "web_url": "https://www.yammer.com/nubo.eu/messages/470839661887488", "group_id": 13114941440, "body": { "parsed": "send a letter to me", "plain": "send a letter to me", "rich": "send a letter to me" }, "thread_id": 470839661887488, "client_type": "O365 Api Auth", "client_url": "https://api.yammer.com", "system_message": false, "direct_message": false, "chat_client_sequence": null, "language": null, "notified_user_ids": [], "privacy": "public", "attachments": [], "liked_by": { "count": 0, "names": [] }, "content_excerpt": "send a letter to me", "group_created_id": 13114941440 }] };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
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

  it('repliedToId must be a number', () => {
    const actual = command.validate({ options: { body: "test", repliedToId: 'nonumber' } });
    assert.notStrictEqual(actual, true);
  });

  it('groupId must be a number', () => {
    const actual = command.validate({ options: { body: "test", groupId: 'nonumber' } });
    assert.notStrictEqual(actual, true);
  });

  it('networkId must be a number', () => {
    const actual = command.validate({ options: { body: "test", networkId: 'nonumber' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if body and repliedToId set', () => {
    const actual = command.validate({ options: { body: "test", repliedToId: 1234122 } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if body and directToUserIds set', () => {
    const actual = command.validate({ options: { body: "test", directToUserIds: 1234125 } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if body and groupId set', () => {
    const actual = command.validate({ options: { body: "test", groupId: 1234123 } });
    assert.strictEqual(actual, true);
  });

  it('fails if no groupId, directToUserId, or repliedToId is specified', () => {
    const actual = command.validate({ options: { body: "test" } });
    assert.notStrictEqual(actual, true);
  });

  it('posts a message', function (done) {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(firstMessage);
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { body: "send a letter to me", groupId: 13114941440, debug: true } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 470839661887488);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('posts a message handling json', function (done) {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(firstMessage);
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { body: "send a letter to me", groupId: 13114941440, debug: true, output: "json" } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 470839661887488);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});