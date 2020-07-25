import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./message-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.YAMMER_MESSAGE_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let firstMessage: any = { messages: [{ "id": 470839661887488, "sender_id": 1496550646, "replied_to_id": null, "created_at": "2019/12/22 17:20:30 +0000", "network_id": 801445, "message_type": "update", "sender_type": "user", "url": "https://www.yammer.com/api/v1/messages/470839661887488", "web_url": "https://www.yammer.com/nubo.eu/messages/470839661887488", "group_id": 13114941440, "body": { "parsed": "send a letter to me", "plain": "send a letter to me", "rich": "send a letter to me" }, "thread_id": 470839661887488, "client_type": "O365 Api Auth", "client_url": "https://api.yammer.com", "system_message": false, "direct_message": false, "chat_client_sequence": null, "language": null, "notified_user_ids": [], "privacy": "public", "attachments": [], "liked_by": { "count": 0, "names": [] }, "content_excerpt": "send a letter to me", "group_created_id": 13114941440 }] };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_MESSAGE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('repliedToId must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { body: "test", repliedToId: 'nonumber' } });
    assert.notStrictEqual(actual, true);
  });

  it('groupId must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { body: "test", groupId: 'nonumber' } });
    assert.notStrictEqual(actual, true);
  });

  it('networkId must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { body: "test", networkId: 'nonumber' } });
    assert.notStrictEqual(actual, true);
  });

  it('has all fields', () => {
    const actual = (command.validate() as CommandValidate)({ options: { body: "test" } });
    assert.strictEqual(actual, true);
  });

  it('posts a message', function (done) {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(firstMessage);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { body: "send a letter to me", debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].id, 470839661887488)
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
    cmdInstance.action({ options: { body: "send a letter to me", debug: true, output: "json" } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].id, 470839661887488)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});