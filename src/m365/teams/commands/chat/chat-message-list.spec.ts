import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./chat-message-list');

describe(commands.CHAT_MESSAGE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CHAT_MESSAGE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'shortBody']);
  });

  it('fails validation if chatId is not specified', () => {
    const actual = command.validate({
      options: {
        debug: false
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the chatId is not valid', () => {
    const actual = command.validate({
      options: {
        chatId: "2da4c29f6d7041"
      }
    });
    assert.notStrictEqual(actual, true);
  });


  it('fails validatation for an incorrect chatId missing leading 19:.', (done) => {
    const actual = command.validate({
      options: {
        chatId: '2da4c29f6d7041eca70b638b43d45437@thread.v2'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation for an incorrect chatId missing trailing @thread.v2 or @unq.gbl.spaces', (done) => {
    const actual = command.validate({
      options: {
        chatId: '19:2da4c29f6d7041eca70b638b43d45437'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
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

  it('validates for a correct input', () => {
    const actual = command.validate({
      options: {
        chatId: "19:2da4c29f6d7041eca70b638b43d45437@thread.v2"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('lists chat messages (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages`) {
        return Promise.resolve({
          "value": [{ "id": "1616964509832", "replyToId": null, "etag": "1616964509832", "messageType": "message", "createdDateTime": "2021-03-28T20:48:29.832Z", "lastModifiedDateTime": "2021-03-28T20:48:29.832Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "text", "content": "Hello world" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615971548136", "replyToId": null, "etag": "1615971548136", "messageType": "message", "createdDateTime": "2021-03-17T08:59:08.136Z", "lastModifiedDateTime": "2021-03-17T08:59:08.136Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "html", "content": "<div><div><div><span><img height=\"63\" src=\"https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages/1615971548136/hostedContents/aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv/$value\" width=\"67\" style=\"vertical-align:bottom; width:67px; height:63px\"></span></div></div></div>" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615943825123", "replyToId": null, "etag": "1615943825123", "messageType": "unknownFutureValue", "createdDateTime": "2021-03-1706:47:05.123Z", "lastModifiedDateTime": "2021-03-1706:47:05.123Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "from": null, "body": { "contentType": "html", "content": "<systemEventMessage/>" }, "attachments": [], "mentions": [], "reactions": [], "eventDetail": { "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail", "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "chatDisplayName": "Graph Members", "initiator": { "application": null, "device": null, "user": { "id": "1fb8890f-423e-4154-8fbf-db6809bc8756", "displayName": null, "userIdentityType": "aadUser" } } } }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: true,
        chatId: "19:2da4c29f6d7041eca70b638b43d45437@thread.v2"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "1616964509832",
            "replyToId": null,
            "etag": "1616964509832",
            "messageType": "message",
            "createdDateTime": "2021-03-28T20:48:29.832Z",
            "lastModifiedDateTime": "2021-03-28T20:48:29.832Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
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
            "body": "Hello world",
            "attachments": [],
            "mentions": [],
            "reactions": [],
            "shortBody": "Hello world"
          },
          {
            "id": "1615971548136",
            "replyToId": null,
            "etag": "1615971548136",
            "messageType": "message",
            "createdDateTime": "2021-03-17T08:59:08.136Z",
            "lastModifiedDateTime": "2021-03-17T08:59:08.136Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
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
            "body": "<div><div><div><span><img height=\"63\" src=\"https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages/1615971548136/hostedContents/aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv/$value\" width=\"67\" style=\"vertical-align:bottom; width:67px; height:63px\"></span></div></div></div>",
            "attachments": [],
            "mentions": [],
            "reactions": [],
            "shortBody": "<div><div><div><span><img height=\"63\" src=\"https:/..."
          },
          {
            "id": "1615943825123",
            "replyToId": null,
            "etag": "1615943825123",
            "messageType": "unknownFutureValue",
            "createdDateTime": "2021-03-1706:47:05.123Z",
            "lastModifiedDateTime": "2021-03-1706:47:05.123Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
            "policyViolation": null,
            "from": null,
            "body": "<systemEventMessage/>",
            "attachments": [],
            "mentions": [],
            "reactions": [],
            "eventDetail": {
              "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail",
              "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
              "chatDisplayName": "Graph Members",
              "initiator": {
                "application": null,
                "device": null,
                "user": {
                  "id": "1fb8890f-423e-4154-8fbf-db6809bc8756",
                  "displayName": null,
                  "userIdentityType": "aadUser"
                }
              }
            },
            "shortBody": "<systemEventMessage/>"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists chat messages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages`) {
        return Promise.resolve({
          value: [{ "id": "1616964509832", "replyToId": null, "etag": "1616964509832", "messageType": "message", "createdDateTime": "2021-03-28T20:48:29.832Z", "lastModifiedDateTime": "2021-03-28T20:48:29.832Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "text", "content": "Hello world" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615971548136", "replyToId": null, "etag": "1615971548136", "messageType": "message", "createdDateTime": "2021-03-17T08:59:08.136Z", "lastModifiedDateTime": "2021-03-17T08:59:08.136Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "html", "content": "<div><div><div><span><img height=\"63\" src=\"https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages/1615971548136/hostedContents/aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv/$value\" width=\"67\" style=\"vertical-align:bottom; width:67px; height:63px\"></span></div></div></div>" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615943825123", "replyToId": null, "etag": "1615943825123", "messageType": "unknownFutureValue", "createdDateTime": "2021-03-1706:47:05.123Z", "lastModifiedDateTime": "2021-03-1706:47:05.123Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "from": null, "body": { "contentType": "html", "content": "<systemEventMessage/>" }, "attachments": [], "mentions": [], "reactions": [], "eventDetail": { "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail", "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "chatDisplayName": "Graph Members", "initiator": { "application": null, "device": null, "user": { "id": "1fb8890f-423e-4154-8fbf-db6809bc8756", "displayName": null, "userIdentityType": "aadUser" } } } }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: false,
        chatId: "19:2da4c29f6d7041eca70b638b43d45437@thread.v2"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "1616964509832",
            "replyToId": null,
            "etag": "1616964509832",
            "messageType": "message",
            "createdDateTime": "2021-03-28T20:48:29.832Z",
            "lastModifiedDateTime": "2021-03-28T20:48:29.832Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
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
            "body": "Hello world",
            "attachments": [],
            "mentions": [],
            "reactions": [],
            "shortBody": "Hello world"
          },
          {
            "id": "1615971548136",
            "replyToId": null,
            "etag": "1615971548136",
            "messageType": "message",
            "createdDateTime": "2021-03-17T08:59:08.136Z",
            "lastModifiedDateTime": "2021-03-17T08:59:08.136Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
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
            "body": "<div><div><div><span><img height=\"63\" src=\"https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages/1615971548136/hostedContents/aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv/$value\" width=\"67\" style=\"vertical-align:bottom; width:67px; height:63px\"></span></div></div></div>",
            "attachments": [],
            "mentions": [],
            "reactions": [],
            "shortBody": "<div><div><div><span><img height=\"63\" src=\"https:/..."
          },
          {
            "id": "1615943825123",
            "replyToId": null,
            "etag": "1615943825123",
            "messageType": "unknownFutureValue",
            "createdDateTime": "2021-03-1706:47:05.123Z",
            "lastModifiedDateTime": "2021-03-1706:47:05.123Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
            "policyViolation": null,
            "from": null,
            "body": "<systemEventMessage/>",
            "attachments": [],
            "mentions": [],
            "reactions": [],
            "eventDetail": {
              "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail",
              "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
              "chatDisplayName": "Graph Members",
              "initiator": {
                "application": null,
                "device": null,
                "user": {
                  "id": "1fb8890f-423e-4154-8fbf-db6809bc8756",
                  "displayName": null,
                  "userIdentityType": "aadUser"
                }
              }
            },
            "shortBody": "<systemEventMessage/>"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('lists chat messages, with truncated shortbody property in text mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages`) {
        return Promise.resolve({
          value: [{ "id": "1616964509832", "replyToId": null, "etag": "1616964509832", "messageType": "message", "createdDateTime": "2021-03-28T20:48:29.832Z", "lastModifiedDateTime": "2021-03-28T20:48:29.832Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "text", "content": "Hello world" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615971548136", "replyToId": null, "etag": "1615971548136", "messageType": "message", "createdDateTime": "2021-03-17T08:59:08.136Z", "lastModifiedDateTime": "2021-03-17T08:59:08.136Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "html", "content": "<div><div><div><span><img height=\"63\" src=\"https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages/1615971548136/hostedContents/aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv/$value\" width=\"67\" style=\"vertical-align:bottom; width:67px; height:63px\"></span></div></div></div>" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615943825123", "replyToId": null, "etag": "1615943825123", "messageType": "unknownFutureValue", "createdDateTime": "2021-03-1706:47:05.123Z", "lastModifiedDateTime": "2021-03-1706:47:05.123Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "from": null, "body": { "contentType": "html", "content": "<systemEventMessage/>" }, "attachments": [], "mentions": [], "reactions": [], "eventDetail": { "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail", "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "chatDisplayName": "Graph Members", "initiator": { "application": null, "device": null, "user": { "id": "1fb8890f-423e-4154-8fbf-db6809bc8756", "displayName": null, "userIdentityType": "aadUser" } } } }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: "text",
        chatId: "19:2da4c29f6d7041eca70b638b43d45437@thread.v2"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "1616964509832",
            "replyToId": null,
            "etag": "1616964509832",
            "messageType": "message",
            "createdDateTime": "2021-03-28T20:48:29.832Z",
            "lastModifiedDateTime": "2021-03-28T20:48:29.832Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
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
            "body": "Hello world",
            "shortBody": "Hello world",
            "attachments": [],
            "mentions": [],
            "reactions": []
          },
          {
            "id": "1615971548136",
            "replyToId": null,
            "etag": "1615971548136",
            "messageType": "message",
            "createdDateTime": "2021-03-17T08:59:08.136Z",
            "lastModifiedDateTime": "2021-03-17T08:59:08.136Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
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
            "body": "<div><div><div><span><img height=\"63\" src=\"https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages/1615971548136/hostedContents/aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv/$value\" width=\"67\" style=\"vertical-align:bottom; width:67px; height:63px\"></span></div></div></div>",
            "shortBody": "<div><div><div><span><img height=\"63\" src=\"https:/...",
            "attachments": [],
            "mentions": [],
            "reactions": []
          },
          {
            "id": "1615943825123",
            "replyToId": null,
            "etag": "1615943825123",
            "messageType": "unknownFutureValue",
            "createdDateTime": "2021-03-1706:47:05.123Z",
            "lastModifiedDateTime": "2021-03-1706:47:05.123Z",
            "lastEditedDateTime": null,
            "deletedDateTime": null,
            "subject": null,
            "summary": null,
            "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
            "importance": "normal",
            "locale": "en-us",
            "webUrl": null,
            "channelIdentity": null,
            "policyViolation": null,
            "from": null,
            "body": "<systemEventMessage/>",
            "shortBody": "<systemEventMessage/>",
            "attachments": [],
            "mentions": [],
            "reactions": [],
            "eventDetail": {
              "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail",
              "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
              "chatDisplayName": "Graph Members",
              "initiator": {
                "application": null,
                "device": null,
                "user": {
                  "id": "1fb8890f-423e-4154-8fbf-db6809bc8756",
                  "displayName": null,
                  "userIdentityType": "aadUser"
                }
              }
            }
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('outputs all data in json output mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages`) {
        return Promise.resolve({
          value: [{ "id": "1616964509832", "replyToId": null, "etag": "1616964509832", "messageType": "message", "createdDateTime": "2021-03-28T20:48:29.832Z", "lastModifiedDateTime": "2021-03-28T20:48:29.832Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "text", "content": "Hello world" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615971548136", "replyToId": null, "etag": "1615971548136", "messageType": "message", "createdDateTime": "2021-03-17T08:59:08.136Z", "lastModifiedDateTime": "2021-03-17T08:59:08.136Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "html", "content": "<div><div><div><span><img height=\"63\" src=\"https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages/1615971548136/hostedContents/aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv/$value\" width=\"67\" style=\"vertical-align:bottom; width:67px; height:63px\"></span></div></div></div>" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615943825123", "replyToId": null, "etag": "1615943825123", "messageType": "unknownFutureValue", "createdDateTime": "2021-03-1706:47:05.123Z", "lastModifiedDateTime": "2021-03-1706:47:05.123Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "from": null, "body": { "contentType": "html", "content": "<systemEventMessage/>" }, "attachments": [], "mentions": [], "reactions": [], "eventDetail": { "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail", "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "chatDisplayName": "Graph Members", "initiator": { "application": null, "device": null, "user": { "id": "1fb8890f-423e-4154-8fbf-db6809bc8756", "displayName": null, "userIdentityType": "aadUser" } } } }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        chatId: "19:2da4c29f6d7041eca70b638b43d45437@thread.v2"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([{ "id": "1616964509832", "replyToId": null, "etag": "1616964509832", "messageType": "message", "createdDateTime": "2021-03-28T20:48:29.832Z", "lastModifiedDateTime": "2021-03-28T20:48:29.832Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "text", "content": "Hello world" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615971548136", "replyToId": null, "etag": "1615971548136", "messageType": "message", "createdDateTime": "2021-03-17T08:59:08.136Z", "lastModifiedDateTime": "2021-03-17T08:59:08.136Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "eventDetail": null, "from": { "application": null, "device": null, "user": { "id": "8ea0e38b-efb3-4757-924a-5f94061cf8c2", "displayName": "Robin Kline", "userIdentityType": "aadUser" } }, "body": { "contentType": "html", "content": "<div><div><div><span><img height=\"63\" src=\"https://graph.microsoft.com/v1.0/chats/19:2da4c29f6d7041eca70b638b43d45437@thread.v2/messages/1615971548136/hostedContents/aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv/$value\" width=\"67\" style=\"vertical-align:bottom; width:67px; height:63px\"></span></div></div></div>" }, "attachments": [], "mentions": [], "reactions": [] }, { "id": "1615943825123", "replyToId": null, "etag": "1615943825123", "messageType": "unknownFutureValue", "createdDateTime": "2021-03-1706:47:05.123Z", "lastModifiedDateTime": "2021-03-1706:47:05.123Z", "lastEditedDateTime": null, "deletedDateTime": null, "subject": null, "summary": null, "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "importance": "normal", "locale": "en-us", "webUrl": null, "channelIdentity": null, "policyViolation": null, "from": null, "body": { "contentType": "html", "content": "<systemEventMessage/>" }, "attachments": [], "mentions": [], "reactions": [], "eventDetail": { "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail", "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2", "chatDisplayName": "Graph Members", "initiator": { "application": null, "device": null, "user": { "id": "1fb8890f-423e-4154-8fbf-db6809bc8756", "displayName": null, "userIdentityType": "aadUser" } } } }]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when listing messages', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        debug: false,
        chatId: "19:2da4c29f6d7041eca70b638b43d45437@thread.v2"
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});