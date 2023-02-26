import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { CommandInfo } from '../../../../cli/CommandInfo';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { Cli } from '../../../../cli/Cli';
import * as userGetCommand from '../../../aad/commands/user/user-get';
import { accessToken } from '../../../../utils/accessToken';
const command: Command = require('./meeting-get');

describe(commands.MEETING_GET, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const userName = 'user@tenant.com';
  const email = 'user@tenant.com';
  const joinUrl = 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_OWIwM2MzNmQtZmY1My00MzM0LWIxMGQtYzkyNzI3OWU5ODMx%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d';
  const meetingResponse = {
    value: [
      {
        "id": "AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAENAABiOC8xvYmdT6G2E_hLMK5kAAIw3TQIAAA=",
        "createdDateTime": "2022-06-26T12:39:34.224055Z",
        "lastModifiedDateTime": "2022-06-26T12:41:36.4357085Z",
        "changeKey": "YjgvMb2JnU+hthPoSzCuZAACMHITIQ==",
        "categories": [],
        "transactionId": null,
        "originalStartTimeZone": "W. Europe Standard Time",
        "originalEndTimeZone": "W. Europe Standard Time",
        "iCalUId": "040000008200E00074C5B7101A82E008000000001AF70ACA5989D801000000000000000010000000048716A892ACAE4DB6CC16097796C401",
        "reminderMinutesBeforeStart": 15,
        "isReminderOn": true,
        "hasAttachments": false,
        "subject": "Test",
        "bodyPreview": "________________________________________________________________________________\r\\\nMicrosoft Teams meeting\r\\\nJoin on your computer or mobile app\r\\\nClick here to join the meeting\r\\\nLearn More | Meeting options\r\\\n___________________________________________",
        "importance": "normal",
        "sensitivity": "normal",
        "isAllDay": false,
        "isCancelled": false,
        "isOrganizer": true,
        "responseRequested": true,
        "seriesMasterId": null,
        "showAs": "busy",
        "type": "singleInstance",
        "webLink": "https://outlook.office365.com/owa/?itemid=AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E%2BhLMK5kAAAAAAENAABiOC8xvYmdT6G2E%2BhLMK5kAAIw3TQIAAA%3D&exvsurl=1&path=/calendar/item",
        "onlineMeetingUrl": null,
        "isOnlineMeeting": true,
        "onlineMeetingProvider": "teamsForBusiness",
        "allowNewTimeProposals": true,
        "occurrenceId": null,
        "isDraft": false,
        "hideAttendees": false,
        "responseStatus": {
          "response": "organizer",
          "time": "0001-01-01T00:00:00Z"
        },
        "body": {
          "contentType": "html",
          "content": "<html>\r\\\n<head>\r\\\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\\\n</head>\r\\\n<body>\r\\\n<div><br>\r\\\n<br>\r\\\n<br>\r\\\n<div style=\"width:100%; height:20px\"><span style=\"white-space:nowrap; color:#5F5F5F; opacity:.36\">________________________________________________________________________________</span>\r\\\n</div>\r\\\n<div class=\"me-email-text\" lang=\"en-US\" style=\"color:#252424; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n<div style=\"margin-top:24px; margin-bottom:20px\"><span style=\"font-size:24px; color:#252424\">Microsoft Teams meeting</span>\r\\\n</div>\r\\\n<div style=\"margin-bottom:20px\">\r\\\n<div style=\"margin-top:0px; margin-bottom:0px; font-weight:bold\"><span style=\"font-size:14px; color:#252424\">Join on your computer or mobile app</span>\r\\\n</div>\r\\\n<a href=\"https://teams.microsoft.com/l/meetup-join/19%3ameeting_OWIwM2MzNmQtZmY1My00MzM0LWIxMGQtYzkyNzI3OWU5ODMx%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d\" class=\"me-email-headline\" style=\"font-size:14px; font-family:'Segoe UI Semibold','Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif; text-decoration:underline; color:#6264a7\">Click\r\\\n here to join the meeting</a> </div>\r\\\n<div style=\"margin-bottom:24px; margin-top:20px\"><a href=\"https://aka.ms/JoinTeamsMeeting\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">Learn More</a>\r\\\n | <a href=\"https://teams.microsoft.com/meetingOptions/?organizerId=fe36f75e-c103-410b-a18a-2bf6df06ac3a&amp;tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d&amp;threadId=19_meeting_OWIwM2MzNmQtZmY1My00MzM0LWIxMGQtYzkyNzI3OWU5ODMx@thread.v2&amp;messageId=0&amp;language=en-US\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\nMeeting options</a> </div>\r\\\n</div>\r\\\n<div style=\"font-size:14px; margin-bottom:4px; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n</div>\r\\\n<div style=\"font-size:12px\"></div>\r\\\n</div>\r\\\n<div style=\"width:100%; height:20px\"><span style=\"white-space:nowrap; color:#5F5F5F; opacity:.36\">________________________________________________________________________________</span>\r\\\n</div>\r\\\n<div></div>\r\\\n</body>\r\\\n</html>\r\\\n"
        },
        "start": {
          "dateTime": "2022-06-26T12:30:00.0000000",
          "timeZone": "UTC"
        },
        "end": {
          "dateTime": "2022-06-26T13:00:00.0000000",
          "timeZone": "UTC"
        },
        "location": {
          "displayName": "",
          "locationType": "default",
          "uniqueIdType": "unknown",
          "address": {},
          "coordinates": {}
        },
        "locations": [],
        "recurrence": null,
        "attendees": [
          {
            "type": "required",
            "status": {
              "response": "none",
              "time": "0001-01-01T00:00:00Z"
            },
            "emailAddress": {
              "name": "User D",
              "address": "userD@outlook.com"
            }
          }
        ],
        "organizer": {
          "emailAddress": {
            "name": "User",
            "address": "user@tenant.com"
          }
        },
        "onlineMeeting": {
          "joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_OWIwM2MzNmQtZmY1My00MzM0LWIxMGQtYzkyNzI3OWU5ODMx%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d"
        }
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
      accessToken.isAppOnlyAccessToken,
      request.get,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.MEETING_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when the userId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        userId: 'foo',
        joinUrl: joinUrl
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified', async () => {
    const actual = await command.validate({
      options: {
        userId: userId,
        joinUrl: joinUrl
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves specific meeting details using userId (debug)', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(joinUrl)}'`) {
        return meetingResponse;
      }

      throw `Invalid request`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        userId: userId,
        joinUrl: joinUrl
      }
    });

    assert(loggerLogSpy.calledWith(meetingResponse.value[0]));
  });

  it('retrieves specific meeting details using userName', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => true);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === userGetCommand) {
        return { "stdout": JSON.stringify({ id: userId }) };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(joinUrl)}'`) {
        return meetingResponse;
      }

      throw `Invalid request`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        userName: userName,
        joinUrl: joinUrl
      }
    });

    assert(loggerLogSpy.calledWith(meetingResponse.value[0]));
  });

  it('retrieves specific meeting details using email', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => true);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === userGetCommand) {
        return { "stdout": JSON.stringify({ id: userId }) };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(joinUrl)}'`) {
        return meetingResponse;
      }

      throw `Invalid request`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        email: email,
        joinUrl: joinUrl
      }
    });

    assert(loggerLogSpy.calledWith(meetingResponse.value[0]));
  });

  it('retrieves specific meeting details using delegated permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(joinUrl)}'`) {
        return meetingResponse;
      }

      throw `Invalid request`;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        joinUrl: joinUrl
      }
    });

    assert(loggerLogSpy.calledWith(meetingResponse.value[0]));
  });

  it('correctly handles error when the meeting with join URL not found', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(joinUrl)}'`) {
        return { value: [] };
      }

      throw `The specified meeting was not found`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        userId: userId,
        joinUrl: joinUrl
      }
    }), new CommandError(`The specified meeting was not found`));
  });

  it('correctly handles error when getting specified meeting details', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => false);

    const errorMessage = 'An error has occurred.';
    sinon.stub(request, 'get').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        joinUrl: joinUrl
      }
    }), new CommandError(errorMessage));
  });

  it('correctly handles error when getting specified meeting details using app only permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => true);

    const errorMessage = `The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`;

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        joinUrl: joinUrl
      }
    }), new CommandError(errorMessage));
  });

  it('correctly handles error when getting specified meeting details using delegated permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => false);
    const errorMessage = `The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`;

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        userId: userId,
        joinUrl: joinUrl
      }
    }), new CommandError(errorMessage));
  });
});
