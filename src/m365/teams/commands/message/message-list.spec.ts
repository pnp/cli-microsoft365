import commands from '../../commands';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./message-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_MESSAGE_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_MESSAGE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if teamId and channelId are not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "fce9e580-8bba-",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validatation for a incorrect channelId missing leading 19:.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation for a incorrect channelId missing trailing @thread.skpye.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread',
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation for since date wrong format', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        since: "2019.12.31"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for since date too far in the past (> 8 months)', () => {
    let d: Date = new Date()
    d.setMonth(d.getMonth() - 9);
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        since: d.toISOString()
      }
    });
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

  it('validates for a correct input', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input (with optional --since param)', () => {
    let d: Date = new Date()
    d.setMonth(d.getMonth() - 7);
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        since: d.toISOString()
      }
    });
    assert.strictEqual(actual, true);
  });

  it('lists messages (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages`) {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/beta/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages?$skiptoken=%2bRID%3avpsQAJ9uAC3rtFEAAADADw%3d%3d%23RT%3a1%23TRC%3a20%23RTD%3auDcxOTplYjMwOTczYjQyYTg0N2EyYTFkZjkyZDkxZTM3Yzc2YUB0aHJlYWQuc2t5cGU7MTUxMTcyMzY2MzY2MA%3d%3d%23FPC%3aAghGAQAAAD8AAIgBAAAAPwAARgEAAAA%2fAAAMAMIzAAwDAAIBAPgBAGoBAAAAPwAACADyBwAwgABmgIgBAAAAPwAAFABTh%2fIEQgDAAGuJAIAhABwA8QJQAA%3d%3d",
          value: [
            {
              "attachments": [],
              "body": {
                "content": "<p>Welcome!</p>",
                "contentType": "html"
              },
              "createdDateTime": "2018-11-15T13:56:40.091Z",
              "deleted": false,
              "etag": "1542290200091",
              "from": {
                "application": {
                  "applicationIdentityType": "bot",
                  "displayName": "POITBot",
                  "id": "d22ece15-e04f-453a-adbd-d1514d2f1abe"
                },
                "conversation": null,
                "device": null,
                "user": null
              },
              "id": "1542290200091",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": null,
              "summary": null
            },
            {
              "attachments": [],
              "body": {
                "content": "hello",
                "contentType": "text"
              },
              "createdDateTime": "2018-11-15T13:20:43.581Z",
              "deleted": false,
              "etag": "1542288043581",
              "from": {
                "application": null,
                "conversation": null,
                "device": null,
                "user": {
                  "displayName": "Balamurugan Kailasam",
                  "id": "065868eb-f08f-4a82-9786-690bc5c38fce",
                  "userIdentityType": "aadUser"
                }
              },
              "id": "1542288043581",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": "",
              "summary": null
            }
          ]
        });
      }
      else if (opts.url === `https://graph.microsoft.com/beta/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages?$skiptoken=%2bRID%3avpsQAJ9uAC3rtFEAAADADw%3d%3d%23RT%3a1%23TRC%3a20%23RTD%3auDcxOTplYjMwOTczYjQyYTg0N2EyYTFkZjkyZDkxZTM3Yzc2YUB0aHJlYWQuc2t5cGU7MTUxMTcyMzY2MzY2MA%3d%3d%23FPC%3aAghGAQAAAD8AAIgBAAAAPwAARgEAAAA%2fAAAMAMIzAAwDAAIBAPgBAGoBAAAAPwAACADyBwAwgABmgIgBAAAAPwAAFABTh%2fIEQgDAAGuJAIAhABwA8QJQAA%3d%3d`) {
        return Promise.resolve({
          value: [
            {
              "attachments": [],
              "body": {
                "content": "Hi...files uploaded",
                "contentType": "html"
              },
              "createdDateTime": "2017-11-26T19:14:23.66Z",
              "deleted": false,
              "etag": "1511723663660",
              "from": {
                "application": null,
                "conversation": null,
                "device": null,
                "user": {
                  "displayName": "orgid:065868eb-f08f-4a82-9786-690bc5c38fce",
                  "id": "065868eb-f08f-4a82-9786-690bc5c38fce",
                  "userIdentityType": "aadUser"
                }
              },
              "id": "1511723663660",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": null,
              "summary": null
            }
          ]
        });
      }
      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "1542290200091",
            "summary": null,
            "body": "<p>Welcome!</p>"
          },
          {
            "id": "1542288043581",
            "summary": null,
            "body": "hello"
          },
          {
            "id": "1511723663660",
            "summary": null,
            "body": "Hi...files uploaded"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists messages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages`) {
        return Promise.resolve({
          value: [
            {
              "attachments": [],
              "body": {
                "content": "<p>Welcome!</p>",
                "contentType": "html"
              },
              "createdDateTime": "2018-11-15T13:56:40.091Z",
              "deleted": false,
              "etag": "1542290200091",
              "from": {
                "application": {
                  "applicationIdentityType": "bot",
                  "displayName": "POITBot",
                  "id": "d22ece15-e04f-453a-adbd-d1514d2f1abe"
                },
                "conversation": null,
                "device": null,
                "user": null
              },
              "id": "1542290200091",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": null,
              "summary": null
            },
            {
              "attachments": [],
              "body": {
                "content": "hello",
                "contentType": "text"
              },
              "createdDateTime": "2018-11-15T13:20:43.581Z",
              "deleted": false,
              "etag": "1542288043581",
              "from": {
                "application": null,
                "conversation": null,
                "device": null,
                "user": {
                  "displayName": "Balamurugan Kailasam",
                  "id": "065868eb-f08f-4a82-9786-690bc5c38fce",
                  "userIdentityType": "aadUser"
                }
              },
              "id": "1542288043581",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": "",
              "summary": null
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "1542290200091",
            "summary": null,
            "body": "<p>Welcome!</p>"
          },
          {
            "id": "1542288043581",
            "summary": null,
            "body": "hello"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists messages since date specified', (done) => {
    const dt: string = new Date().toISOString()
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages/delta?$filter=lastModifiedDateTime gt ${dt}`) {
        return Promise.resolve({
          value: [
            {
              "attachments": [],
              "body": {
                "content": "<p>Welcome!</p>",
                "contentType": "html"
              },
              "createdDateTime": "2018-11-15T13:56:40.091Z",
              "deleted": false,
              "etag": "1542290200091",
              "from": {
                "application": {
                  "applicationIdentityType": "bot",
                  "displayName": "POITBot",
                  "id": "d22ece15-e04f-453a-adbd-d1514d2f1abe"
                },
                "conversation": null,
                "device": null,
                "user": null
              },
              "id": "1542290200091",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": null,
              "summary": null
            },
            {
              "attachments": [],
              "body": {
                "content": "hello",
                "contentType": "text"
              },
              "createdDateTime": "2018-11-15T13:20:43.581Z",
              "deleted": false,
              "etag": "1542288043581",
              "from": {
                "application": null,
                "conversation": null,
                "device": null,
                "user": {
                  "displayName": "Balamurugan Kailasam",
                  "id": "065868eb-f08f-4a82-9786-690bc5c38fce",
                  "userIdentityType": "aadUser"
                }
              },
              "id": "1542288043581",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": "",
              "summary": null
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        since: dt 
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "1542290200091",
            "summary": null,
            "body": "<p>Welcome!</p>"
          },
          {
            "id": "1542288043581",
            "summary": null,
            "body": "hello"
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
      if (opts.url === `https://graph.microsoft.com/beta/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages`) {
        return Promise.resolve({
          value: [
            {
              "attachments": [],
              "body": {
                "content": "<p>Welcome!</p>",
                "contentType": "html"
              },
              "createdDateTime": "2018-11-15T13:56:40.091Z",
              "deleted": false,
              "etag": "1542290200091",
              "from": {
                "application": {
                  "applicationIdentityType": "bot",
                  "displayName": "POITBot",
                  "id": "d22ece15-e04f-453a-adbd-d1514d2f1abe"
                },
                "conversation": null,
                "device": null,
                "user": null
              },
              "id": "1542290200091",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": null,
              "summary": null
            },
            {
              "attachments": [],
              "body": {
                "content": "hello",
                "contentType": "text"
              },
              "createdDateTime": "2018-11-15T13:20:43.581Z",
              "deleted": false,
              "etag": "1542288043581",
              "from": {
                "application": null,
                "conversation": null,
                "device": null,
                "user": {
                  "displayName": "Balamurugan Kailasam",
                  "id": "065868eb-f08f-4a82-9786-690bc5c38fce",
                  "userIdentityType": "aadUser"
                }
              },
              "id": "1542288043581",
              "importance": "normal",
              "lastModifiedDateTime": null,
              "locale": "en-us",
              "mentions": [],
              "messageType": "message",
              "policyViolation": null,
              "reactions": [],
              "replyToId": null,
              "subject": "",
              "summary": null
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        output: 'json',
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "attachments": [],
            "body": {
              "content": "<p>Welcome!</p>",
              "contentType": "html"
            },
            "createdDateTime": "2018-11-15T13:56:40.091Z",
            "deleted": false,
            "etag": "1542290200091",
            "from": {
              "application": {
                "applicationIdentityType": "bot",
                "displayName": "POITBot",
                "id": "d22ece15-e04f-453a-adbd-d1514d2f1abe"
              },
              "conversation": null,
              "device": null,
              "user": null
            },
            "id": "1542290200091",
            "importance": "normal",
            "lastModifiedDateTime": null,
            "locale": "en-us",
            "mentions": [],
            "messageType": "message",
            "policyViolation": null,
            "reactions": [],
            "replyToId": null,
            "subject": null,
            "summary": null
          },
          {
            "attachments": [],
            "body": {
              "content": "hello",
              "contentType": "text"
            },
            "createdDateTime": "2018-11-15T13:20:43.581Z",
            "deleted": false,
            "etag": "1542288043581",
            "from": {
              "application": null,
              "conversation": null,
              "device": null,
              "user": {
                "displayName": "Balamurugan Kailasam",
                "id": "065868eb-f08f-4a82-9786-690bc5c38fce",
                "userIdentityType": "aadUser"
              }
            },
            "id": "1542288043581",
            "importance": "normal",
            "lastModifiedDateTime": null,
            "locale": "en-us",
            "mentions": [],
            "messageType": "message",
            "policyViolation": null,
            "reactions": [],
            "replyToId": null,
            "subject": "",
            "summary": null
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when listing messages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({
      options: {
        debug: false,
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    }, (err?: any) => {
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