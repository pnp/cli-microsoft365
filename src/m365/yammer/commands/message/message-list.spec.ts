import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./message-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.YAMMER_MESSAGE_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let firstMessageBatch: any = {
    messages: [{ "sender_id": 1496550646, "replied_to_id": 1496550647, "id": 10123190123123, "thread_id": "", group_id: 11231123123, created_at: "2019/09/09 07:53:18 +0000", "content_excerpt": "message1", "body": { "plain": "message1 message is longer than 25 chars. Just for testing shortening" } },
    { "sender_id": 1496550640, "replied_to_id": "", "id": 10123190123124, "thread_id": "", group_id: "", created_at: "2019/09/08 07:53:18 +0000", "content_excerpt": "message2", "body": { "plain": "message2" } },
    { "sender_id": 1496550610, "replied_to_id": "", "id": 10123190123125, "thread_id": "", group_id: "", created_at: "2019/09/07 07:53:18 +0000", "content_excerpt": "message3", "body": { "plain": "message3" } },
    { "sender_id": 1496550630, "replied_to_id": "", "id": 10123190123126, "thread_id": "", group_id: 1123121, created_at: "2019/09/06 07:53:18 +0000", "content_excerpt": "message4", "body": { "plain": "message4" } },
    { "sender_id": 1496550646, "replied_to_id": "", "id": 10123190123127, "thread_id": "", group_id: 1123121, created_at: "2019/09/05 07:53:18 +0000", "content_excerpt": "message5", "body": { "plain": "message5" } }],
    meta: {
      older_available: true
    }
  };
  let secondMessageBatch: any = {
    messages: [{ "sender_id": 1496550646, "replied_to_id": 1496550647, "id": 10123190123130, "thread_id": "", group_id: 11231123123, created_at: "2019/09/04 07:53:18 +0000", "content_excerpt": "message6", "body": { "plain": "message6" } },
    { "sender_id": 1496550640, "replied_to_id": "", "id": 10123190123131, "thread_id": "", group_id: "", created_at: "2019/09/03 07:53:18 +0000", "content_excerpt": "message7", "body": { "plain": "message7" } }],
    meta: {
      older_available: false
    }
  };

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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_MESSAGE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
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

  it('passes validation without parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { limit: 10 } });
    assert.strictEqual(actual, true);
  });

  it('limit must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { limit: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('olderThanId must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { olderThanId: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('groupId must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { groupId: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('threadId must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { threadId: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('you are not allowed to use groupId and threadId at the same time', () => {
    const actual = (command.validate() as CommandValidate)({ options: { groupId: 123, threadId: 123 } });
    assert.notStrictEqual(actual, true);
  });

  it('you cannot specify the feedType with groupId or threadId at the same time', () => {
    const actual = (command.validate() as CommandValidate)({ options: { feedType: 'All', threadId: 123 } });
    assert.notStrictEqual(actual, true);
  });

  it('Fails in case FeedType is not correct', () => {
    const actual = (command.validate() as CommandValidate)({ options: { feedType: 'WrongValue' } });
    assert.notStrictEqual(actual, true);
  });

  it('you are not allowed to use groupId and threadId and feedType at the same time', () => {
    const actual = (command.validate() as CommandValidate)({ options: { feedType: 'Private', groupId: 123, threadId: 123 } });
    assert.notStrictEqual(actual, true);
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

  it('returns messages without more results', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from top feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/algo.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { feedType: 'Top' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from my feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/my_feed.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { feedType: 'My' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from following feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/following.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { feedType: 'Following' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from sent feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/sent.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { feedType: 'Sent' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from private feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/private.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { feedType: 'Private' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from received feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/received.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { feedType: 'Received' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from all feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { feedType: 'All' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from the group feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/in_group/123123.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { groupId: 123123 } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns messages from thread feed', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/in_thread/123123.json') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { threadId: 123123 } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all messages', (done) => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (i++ === 0) {
        return Promise.resolve(firstMessageBatch);
      }
      else {
        return Promise.resolve(secondMessageBatch);
      }
    });
    cmdInstance.action({ options: { output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].length, 7);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns message with a specific limit', (done) => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (i++ === 0) {
        return Promise.resolve(firstMessageBatch);
      }
      else {
        return Promise.resolve(secondMessageBatch);
      }
    });
    cmdInstance.action({ options: { limit: 6, output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].length, 6);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error in loop', (done) => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake((opts) => {
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
    cmdInstance.action({ options: { output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles correct parameters older than', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json?older_than=10123190123128') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { olderThanId: 10123190123128, output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles correct parameters older than and threaded', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json?older_than=10123190123128&threaded=true') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { olderThanId: 10123190123128, threaded: true, output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles correct parameters with threaded', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages.json?threaded=true') {
        return Promise.resolve(secondMessageBatch);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { threaded: true, output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 10123190123130)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});