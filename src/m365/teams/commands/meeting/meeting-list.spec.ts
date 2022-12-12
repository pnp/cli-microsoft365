import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth, { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { CommandInfo } from '../../../../cli/CommandInfo';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { Cli } from '../../../../cli/Cli';
import * as userGetCommand from '../../../aad/commands/user/user-get';
const command: Command = require('./meeting-list');

describe(commands.MEETING_LIST, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const startDateTime = '2022-01-01';
  const endDateTime = '2022-12-31';
  const userName = 'user@tenant.com';
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
      },
      {
        "id": "AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAENAABiOC8xvYmdT6G2E_hLMK5kAAH8dhmuAAA=",
        "createdDateTime": "2022-04-08T11:48:22.2527866Z",
        "lastModifiedDateTime": "2022-04-08T11:50:24.1356845Z",
        "changeKey": "YjgvMb2JnU+hthPoSzCuZAAB/B2ICg==",
        "categories": [],
        "transactionId": null,
        "originalStartTimeZone": "Romance Standard Time",
        "originalEndTimeZone": "Romance Standard Time",
        "iCalUId": "040000008200E00074C5B7101A82E00800000000A87B618C3E4BD8010000000000000000100000006D28750A6361354E9076FFD0D4C5452E",
        "reminderMinutesBeforeStart": 15,
        "isReminderOn": true,
        "hasAttachments": false,
        "subject": "Test",
        "bodyPreview": "________________________________________________________________________________\r\\\nMicrosoft Teams-vergadering\r\\\nDeelnemen op uw computer of via de mobiele app\r\\\nKlik hier om deel te nemen aan de vergadering\r\\\nMeer informatie | Opties voor vergadering",
        "importance": "normal",
        "sensitivity": "normal",
        "isAllDay": false,
        "isCancelled": false,
        "isOrganizer": true,
        "responseRequested": true,
        "seriesMasterId": null,
        "showAs": "busy",
        "type": "singleInstance",
        "webLink": "https://outlook.office365.com/owa/?itemid=AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E%2BhLMK5kAAAAAAENAABiOC8xvYmdT6G2E%2BhLMK5kAAH8dhmuAAA%3D&exvsurl=1&path=/calendar/item",
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
          "content": "<html>\r\\\n<head>\r\\\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\\\n</head>\r\\\n<body>\r\\\n<div><br>\r\\\n<br>\r\\\n<br>\r\\\n<div style=\"width:100%; height:20px\"><span style=\"white-space:nowrap; color:#5F5F5F; opacity:.36\">________________________________________________________________________________</span>\r\\\n</div>\r\\\n<div class=\"me-email-text\" lang=\"nl-NL\" style=\"color:#252424; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n<div style=\"margin-top:24px; margin-bottom:20px\"><span style=\"font-size:24px; color:#252424\">Microsoft Teams-vergadering</span>\r\\\n</div>\r\\\n<div style=\"margin-bottom:20px\">\r\\\n<div style=\"margin-top:0px; margin-bottom:0px; font-weight:bold\"><span style=\"font-size:14px; color:#252424\">Deelnemen op uw computer of via de mobiele app</span>\r\\\n</div>\r\\\n<a href=\"https://teams.microsoft.com/l/meetup-join/19%3ameeting_MjM1ZDM1ZjYtNTgwOC00MWM4LThiYWItNmZhNmM3MTJjZGZm%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d\" class=\"me-email-headline\" style=\"font-size:14px; font-family:'Segoe UI Semibold','Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif; text-decoration:underline; color:#6264a7\">Klik\r\\\n hier om deel te nemen aan de vergadering</a> </div>\r\\\n<div style=\"margin-bottom:24px; margin-top:20px\"><a href=\"https://aka.ms/JoinTeamsMeeting\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">Meer informatie</a>\r\\\n | <a href=\"https://teams.microsoft.com/meetingOptions/?organizerId=fe36f75e-c103-410b-a18a-2bf6df06ac3a&amp;tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d&amp;threadId=19_meeting_MjM1ZDM1ZjYtNTgwOC00MWM4LThiYWItNmZhNmM3MTJjZGZm@thread.v2&amp;messageId=0&amp;language=nl-NL\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\nOpties voor vergadering</a> </div>\r\\\n</div>\r\\\n<div style=\"font-size:14px; margin-bottom:4px; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n</div>\r\\\n<div style=\"font-size:12px\"></div>\r\\\n</div>\r\\\n<div style=\"width:100%; height:20px\"><span style=\"white-space:nowrap; color:#5F5F5F; opacity:.36\">________________________________________________________________________________</span>\r\\\n</div>\r\\\n<div></div>\r\\\n</body>\r\\\n</html>\r\\\n"
        },
        "start": {
          "dateTime": "2022-04-08T11:30:00.0000000",
          "timeZone": "UTC"
        },
        "end": {
          "dateTime": "2022-04-08T12:00:00.0000000",
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
              "name": "User A",
              "address": "userA@tenant.com"
            }
          }
        ],
        "organizer": {
          "emailAddress": {
            "name": "User B",
            "address": "user@tenant.com"
          }
        },
        "onlineMeeting": {
          "joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_MjM1ZDM1ZjYtNTgwOC00MWM4LThiYWItNmZhNmM3MTJjZGZm%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d"
        }
      },
      {
        "id": "AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAENAABiOC8xvYmdT6G2E_hLMK5kAAHxtR_EAAA=",
        "createdDateTime": "2022-03-23T14:41:00.1950925Z",
        "lastModifiedDateTime": "2022-03-23T14:43:02.1403526Z",
        "changeKey": "YjgvMb2JnU+hthPoSzCuZAAB8WHbHA==",
        "categories": [],
        "transactionId": "2f831e09-5507-24ba-2352-bc29160933ef",
        "originalStartTimeZone": "Aleutian Standard Time",
        "originalEndTimeZone": "Aleutian Standard Time",
        "iCalUId": "040000008200E00074C5B7101A82E0080000000095AA9303C43ED801000000000000000010000000EDB19B20BAF3C548841220C2102492CB",
        "reminderMinutesBeforeStart": 15,
        "isReminderOn": true,
        "hasAttachments": false,
        "subject": "Online meeting test",
        "bodyPreview": "________________________________________________________________________________\r\\\nMicrosoft Teams meeting\r\\\nJoin on your computer or mobile app\r\\\nClick here to join the meeting\r\\\nLearn More | Meeting options\r\\\n_______________________________________________",
        "importance": "normal",
        "sensitivity": "normal",
        "isAllDay": false,
        "isCancelled": false,
        "isOrganizer": true,
        "responseRequested": true,
        "seriesMasterId": null,
        "showAs": "busy",
        "type": "singleInstance",
        "webLink": "https://outlook.office365.com/owa/?itemid=AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E%2BhLMK5kAAAAAAENAABiOC8xvYmdT6G2E%2BhLMK5kAAHxtR%2BEAAA%3D&exvsurl=1&path=/calendar/item",
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
          "content": "<html>\r\\\n<head>\r\\\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\\\n</head>\r\\\n<body>\r\\\n<br>\r\\\n<div style=\"width:100%; height:20px\"><span style=\"white-space:nowrap; color:#5F5F5F; opacity:.36\">________________________________________________________________________________</span>\r\\\n</div>\r\\\n<div class=\"me-email-text\" lang=\"en-US\" style=\"color:#252424; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n<div style=\"margin-top:24px; margin-bottom:20px\"><span style=\"font-size:24px; color:#252424\">Microsoft Teams meeting</span>\r\\\n</div>\r\\\n<div style=\"margin-bottom:20px\">\r\\\n<div style=\"margin-top:0px; margin-bottom:0px; font-weight:bold\"><span style=\"font-size:14px; color:#252424\">Join on your computer or mobile app</span>\r\\\n</div>\r\\\n<a href=\"https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZmIxNmI2MzItMGE0MC00NmYwLWIzNGItYzcwMWJiMmQ3MTY0%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d\" class=\"me-email-headline\" style=\"font-size:14px; font-family:'Segoe UI Semibold','Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif; text-decoration:underline; color:#6264a7\">Click\r\\\n here to join the meeting</a> </div>\r\\\n<div style=\"margin-bottom:24px; margin-top:20px\"><a href=\"https://aka.ms/JoinTeamsMeeting\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">Learn More</a>\r\\\n | <a href=\"https://teams.microsoft.com/meetingOptions/?organizerId=fe36f75e-c103-410b-a18a-2bf6df06ac3a&amp;tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d&amp;threadId=19_meeting_ZmIxNmI2MzItMGE0MC00NmYwLWIzNGItYzcwMWJiMmQ3MTY0@thread.v2&amp;messageId=0&amp;language=en-US\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\nMeeting options</a> </div>\r\\\n</div>\r\\\n<div style=\"font-size:14px; margin-bottom:4px; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n</div>\r\\\n<div style=\"font-size:12px\"></div>\r\\\n<div></div>\r\\\n<div style=\"width:100%; height:20px\"><span style=\"white-space:nowrap; color:#5F5F5F; opacity:.36\">________________________________________________________________________________</span>\r\\\n</div>\r\\\n</body>\r\\\n</html>\r\\\n"
        },
        "start": {
          "dateTime": "2022-03-15T05:00:00.0000000",
          "timeZone": "UTC"
        },
        "end": {
          "dateTime": "2022-03-15T05:30:00.0000000",
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
              "name": "Joni Sherman",
              "address": "JoniS@tenant.com"
            }
          }
        ],
        "organizer": {
          "emailAddress": {
            "name": "User B",
            "address": "user@tenant.com"
          }
        },
        "onlineMeeting": {
          "joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZmIxNmI2MzItMGE0MC00NmYwLWIzNGItYzcwMWJiMmQ3MTY0%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d"
        }
      }
    ]
  };
  const meetingResponseText: any = [
    {
      "subject": "Test",
      "start": "2022-06-26T12:30:00.0000000",
      "end": "2022-06-26T13:00:00.0000000"
    },
    {
      "subject": "Test",
      "start": "2022-04-08T11:30:00.0000000",
      "end": "2022-04-08T12:00:00.0000000"
    },
    {
      "subject": "Online meeting test",
      "start": "2022-03-15T05:00:00.0000000",
      "end": "2022-03-15T05:30:00.0000000"
    }
  ];
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      Auth.isAppOnlyAuth,
      request.get,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.MEETING_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['subject', 'start', 'end']);
  });

  it('lists messages using application permissions for a specific userName and specifying only startDateTime', async () => {
    sinon.stub(Auth, 'isAppOnlyAuth').callsFake(() => true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/events?$filter=start/dateTime ge '${startDateTime}'`) {
        return meetingResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, startDateTime: startDateTime } });
    assert(loggerLogSpy.calledWith(meetingResponse.value));
  });

  it('lists messages using application permissions for a specific userId with a pretty output and specifying both startDateTime and endDateTime', async () => {
    sinon.stub(Auth, 'isAppOnlyAuth').callsFake(() => true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events?$filter=start/dateTime ge '${startDateTime}' and end/dateTime le '${endDateTime}'`) {
        return meetingResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId, startDateTime: startDateTime, endDateTime: endDateTime, output: 'text' } });
    assert(loggerLogSpy.calledWith(meetingResponseText));
  });

  it('lists messages using application permissions for a specific user retrieved by email and specifying all other possible options', async () => {
    sinon.stub(Auth, 'isAppOnlyAuth').callsFake(() => true);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === userGetCommand) {
        return { "stdout": JSON.stringify({ id: userId }) };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events?$filter=start/dateTime ge '${startDateTime}' and end/dateTime le '${endDateTime}'`) {
        return meetingResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { email: userName, startDateTime: startDateTime, endDateTime: endDateTime, output: 'text' } });
    assert(loggerLogSpy.calledWith(meetingResponseText));
  });

  it('lists messages using delegated permissions specifying both startDateTime and only retrieving the events that the user is organizer from', async () => {
    sinon.stub(Auth, 'isAppOnlyAuth').callsFake(() => false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events?$filter=start/dateTime ge '${startDateTime}' and isOrganizer eq true`) {
        return meetingResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, output: 'json', startDateTime: startDateTime, isOrganizer: true } });
    assert(loggerLogSpy.calledWith(meetingResponse.value));
  });

  it('correctly handles error when listing events', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('completes validation when the startDateTime is a valid ISODateTime, endDateTime is a valid ISODateTime and userId is a valid Guid', async () => {
    const actual = await command.validate({ options: { verbose: true, startDateTime: startDateTime, endDateTime: endDateTime, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the startDateTime is not a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { verbose: true, startDateTime: 'foo', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userId is not a valid guid', async () => {
    const actual = await command.validate({ options: { verbose: true, startDateTime: startDateTime, userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the endDateTime is not a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { verbose: true, startDateTime: startDateTime, endDateTime: 'foo', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('throws an error when the userName, userId or email is not filled in when signed in using app-only authentication', async () => {
    sinon.stub(Auth, 'isAppOnlyAuth').callsFake(() => true);

    await assert.rejects(command.action(logger, { options: { verbose: true, startDateTime: '2022-04-04' } } as any),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`));
  });

  it('throws an error when the userName is filled in when signed in using delegated authentication', async () => {
    sinon.stub(Auth, 'isAppOnlyAuth').callsFake(() => false);

    await assert.rejects(command.action(logger, { options: { verbose: true, startDateTime: '2022-04-04', userName: userName } } as any),
      new CommandError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`));
  });
});
