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
import command from './message-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.MESSAGE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'summary', 'body']);
  });

  it('fails validation if teamId and channelId are not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for a incorrect channelId missing leading 19:.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for a incorrect channelId missing trailing @thread.skype.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for since date wrong format', async () => {
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        since: "2019.12.31"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for since date too far in the past (> 8 months)', async () => {
    const d: Date = new Date();
    d.setMonth(d.getMonth() - 9);
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        since: d.toISOString()
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input', async () => {
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input (with optional --since param)', async () => {
    const d: Date = new Date();
    d.setMonth(d.getMonth() - 7);
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        since: d.toISOString()
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists messages (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages`) {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages?$skiptoken=%2bRID%3avpsQAJ9uAC3rtFEAAADADw%3d%3d%23RT%3a1%23TRC%3a20%23RTD%3auDcxOTplYjMwOTczYjQyYTg0N2EyYTFkZjkyZDkxZTM3Yzc2YUB0aHJlYWQuc2t5cGU7MTUxMTcyMzY2MzY2MA%3d%3d%23FPC%3aAghGAQAAAD8AAIgBAAAAPwAARgEAAAA%2fAAAMAMIzAAwDAAIBAPgBAGoBAAAAPwAACADyBwAwgABmgIgBAAAAPwAAFABTh%2fIEQgDAAGuJAIAhABwA8QJQAA%3d%3d",
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
        };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages?$skiptoken=%2bRID%3avpsQAJ9uAC3rtFEAAADADw%3d%3d%23RT%3a1%23TRC%3a20%23RTD%3auDcxOTplYjMwOTczYjQyYTg0N2EyYTFkZjkyZDkxZTM3Yzc2YUB0aHJlYWQuc2t5cGU7MTUxMTcyMzY2MzY2MA%3d%3d%23FPC%3aAghGAQAAAD8AAIgBAAAAPwAARgEAAAA%2fAAAMAMIzAAwDAAIBAPgBAGoBAAAAPwAACADyBwAwgABmgIgBAAAAPwAAFABTh%2fIEQgDAAGuJAIAhABwA8QJQAA%3d%3d`) {
        return {
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
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    });
    assert(loggerLogSpy.calledWith([{
      "attachments": [],
      "body": "<p>Welcome!</p>",
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
      "body": "hello",
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
    },
    {
      "attachments": [],
      "body": "Hi...files uploaded",
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
    }]));
  });

  it('lists messages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    });
    assert(loggerLogSpy.calledWith([{
      "attachments": [],
      "body": "<p>Welcome!</p>",
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
      "body": "hello",
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
    }]));
  });

  it('lists messages since date specified', async () => {
    const dt: string = new Date().toISOString();
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages/delta?$filter=lastModifiedDateTime gt ${dt}`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        since: dt
      }
    });
    assert(loggerLogSpy.calledWith([{
      "attachments": [],
      "body": "<p>Welcome!</p>",
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
      "body": "hello",
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
    }]));
  });

  it('outputs all data in json output mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/fce9e580-8bba-4638-ab5c-ab40016651e3/channels/19:eb30973b42a847a2a1df92d91e37c76a@thread.skype/messages`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    });
    assert(loggerLogSpy.calledWith([
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
  });

  it('correctly handles error when listing messages', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };

    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    } as any), new CommandError('An error has occurred'));
  });
});
