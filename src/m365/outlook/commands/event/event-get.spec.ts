import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './event-get.js';
import { calendar } from '../../../../utils/calendar.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.EVENT_GET, () => {
  const id = "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAADMN-7V4K8g0q_adetip1DygcAxMBBaLl1lk_dAn8KkjfXKQAAAgENAAAAxMBBaLl1lk_dAn8KkjfXKQAGMVCCQQAAAA==";
  const userId = "9bd29c6c-181e-41f5-a1b6-bc30bbf652d3";
  const userName = "john.doe@contoso.com";
  const calendarId = "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAAAAAEGAADEwEFouXWWT50CfwqSN9cpAAAkuACjAAA=";
  const calendarName = "Calendar";

  const eventResponse = {
    "id": "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAADMN-7V4K8g0q_adetip1DygcAxMBBaLl1lk_dAn8KkjfXKQAAAgENAAAAxMBBaLl1lk_dAn8KkjfXKQAGMVCCQQAAAA==",
    "createdDateTime": "2026-04-04T11:03:22.881996Z",
    "lastModifiedDateTime": "2026-04-04T11:05:26.2216557Z",
    "changeKey": "xMBBaLl1lk+dAn8KkjfXKQAGLmp8jA==",
    "categories": [],
    "transactionId": "localevent:93639269-b1b2-d604-5170-283b0e470da5",
    "originalStartTimeZone": "UTC",
    "originalEndTimeZone": "UTC",
    "iCalUId": "040000008200E00074C5B7101A82E0080000000051EE49A722C4DC0100000000000000001000000065853ABD35D4FE438112E0B9CF451ABF",
    "uid": "040000008200E00074C5B7101A82E0080000000051EE49A722C4DC0100000000000000001000000065853ABD35D4FE438112E0B9CF451ABF",
    "reminderMinutesBeforeStart": 15,
    "isReminderOn": true,
    "hasAttachments": false,
    "subject": "New Product Regulations Touchpoint",
    "bodyPreview": "New Product Regulations Strategy Online Touchpoint Meeting\r\\\n\r\\\nYou're receiving this message because you're a member of the Engineering group. If you don't want to receive any messages or events from this group, stop following it in your inbox.\r\\\n\r\\\n________",
    "importance": "normal",
    "sensitivity": "normal",
    "isAllDay": false,
    "isCancelled": false,
    "isOrganizer": true,
    "responseRequested": true,
    "seriesMasterId": null,
    "showAs": "busy",
    "type": "singleInstance",
    "webLink": "https://outlook.office365.com/owa/?itemid=AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAADMN%2F7V4K8g0q%2Badetip1DygcAxMBBaLl1lk%2BdAn8KkjfXKQAAAgENAAAAxMBBaLl1lk%2BdAn8KkjfXKQAGMVCCQQAAAA%3D%3D&exvsurl=1&path=/calendar/item",
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
      "content": "<html>\r\\\n<head>\r\\\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\\\n</head>\r\\\n<body>\r\\\n<div style=\"font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)\">\r\\\nNew Product Regulations Strategy Online Touchpoint Meeting</div>\r\\\n<div style=\"font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)\">\r\\\n<br>\r\\\n</div>\r\\\n<div style=\"font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)\">\r\\\nYou're receiving this message because you're a member of the Engineering group. If you don't want to receive any messages or events from this group, stop following it in&nbsp;your inbox.</div>\r\\\n<br>\r\\\n<div class=\"me-email-text\" lang=\"en-US\" style=\"max-width:1024px; color:#242424; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n<div aria-hidden=\"true\" style=\"margin-bottom:24px; overflow:hidden; white-space:nowrap\">\r\\\n________________________________________________________________________________</div>\r\\\n<div style=\"margin-bottom:12px\"><span class=\"me-email-text\" style=\"font-size:20px; color:#242424; font-weight:600\">Microsoft Teams meeting</span>\r\\\n</div>\r\\\n<div style=\"margin-bottom:6px\"><span class=\"me-email-text\" style=\"font-size:20px; color:#242424; font-weight:600\">Join:\r\\\n</span><a href=\"https://teams.microsoft.com/meet/48803137263631?p=YXe9K6OhVD94VIC23M\" id=\"meet_invite_block.action.join_link\" title=\"Meeting join\" aria-label=\"Meeting join\" class=\"me-email-link\" style=\"font-size:20px; text-decoration:underline; color:#5B5FC7\">https://teams.microsoft.com/meet/48803137263631?p=YXe9K6OhVD94VIC23M</a>\r\\\n</div>\r\\\n<div style=\"margin-bottom:6px\"><span class=\"me-email-text-secondary\" style=\"font-size:14px; color:#616161\">Meeting ID:\r\\\n</span><span class=\"me-email-text\" style=\"font-size:14px; color:#242424\">488 031 372 636 31</span>\r\\\n</div>\r\\\n<div style=\"margin-bottom:32px\"><span class=\"me-email-text-secondary\" style=\"font-size:14px; color:#616161\">Passcode:\r\\\n</span><span class=\"me-email-text\" style=\"font-size:14px; color:#242424\">uN2Np6PN</span>\r\\\n</div>\r\\\n<div style=\"margin-bottom:12px; max-width:1024px\">\r\\\n<hr style=\"border:0; background:#616161; height:1px\">\r\\\n</div>\r\\\n<div style=\"margin-bottom:24px\"><a href=\"https://aka.ms/JoinTeamsMeeting?omkt=en-US\" id=\"meet_invite_block.action.help\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#5B5FC7\">Need help?</a>\r\\\n<span style=\"color:#616161\">|</span> <a href=\"https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZjE4ZGNmODktODg3ZS00MTRjLTg4ZmMtZWMzMjBkZTE5YjBl%40thread.v2/0?context=%7b%22Tid%22%3a%22f2c94a41-d33d-4b60-bb3d-0bed4cdf9855%22%2c%22Oid%22%3a%229bd29c6c-181e-41f5-a1b6-bc30bbf652d3%22%7d\" id=\"meet_invite_block.action.join_link_compatibility\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#5B5FC7\">\r\\\nSystem reference</a> </div>\r\\\n<div><span class=\"me-email-text-secondary\" style=\"font-size:14px; color:#616161\">For organizers:\r\\\n</span><a href=\"https://teams.microsoft.com/meetingOptions/?organizerId=9bd29c6c-181e-41f5-a1b6-bc30bbf652d3&amp;tenantId=f2c94a41-d33d-4b60-bb3d-0bed4cdf9855&amp;threadId=19_meeting_ZjE4ZGNmODktODg3ZS00MTRjLTg4ZmMtZWMzMjBkZTE5YjBl@thread.v2&amp;messageId=0&amp;language=en-US\" id=\"meet_invite_block.action.organizer_meet_options\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#5B5FC7\">Meeting\r\\\n options</a> </div>\r\\\n<div style=\"margin-top:24px; margin-bottom:6px\"></div>\r\\\n<div style=\"margin-bottom:24px\"></div>\r\\\n<div aria-hidden=\"true\" style=\"margin-bottom:24px; overflow:hidden; white-space:nowrap\">\r\\\n________________________________________________________________________________</div>\r\\\n</div>\r\\\n</body>\r\\\n</html>\r\\\n"
    },
    "start": {
      "dateTime": "2026-04-04T11:30:00.0000000",
      "timeZone": "UTC"
    },
    "end": {
      "dateTime": "2026-04-04T12:00:00.0000000",
      "timeZone": "UTC"
    },
    "location": {
      "displayName": "Microsoft Teams Meeting",
      "locationType": "default",
      "uniqueId": "Microsoft Teams Meeting",
      "uniqueIdType": "private"
    },
    "locations": [
      {
        "displayName": "Microsoft Teams Meeting",
        "locationType": "default",
        "uniqueId": "Microsoft Teams Meeting",
        "uniqueIdType": "private"
      }
    ],
    "recurrence": null,
    "attendees": [
      {
        "type": "required",
        "status": {
          "response": "none",
          "time": "0001-01-01T00:00:00Z"
        },
        "emailAddress": {
          "name": "Debra Berger",
          "address": "debraB@contoso.com"
        }
      }
    ],
    "organizer": {
      "emailAddress": {
        "name": "John Doe",
        "address": "john.doe@contoso.com"
      }
    },
    "onlineMeeting": {
      "joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZjE4ZGNmODktODg3ZS00MTRjLTg4ZmMtZWMzMjBkZTE5YjBl%40thread.v2/0?context=%7b%22Tid%22%3a%22f2c94a41-d33d-4b60-bb3d-0bed4cdf9855%22%2c%22Oid%22%3a%229bd29c6c-181e-41f5-a1b6-bc30bbf652d3%22%7d"
    }
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
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
      request.get,
      calendar.getUserCalendarByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EVENT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation with userId', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, userId: userId, calendarId: calendarId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with userName', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, userName: userName, calendarName: calendarName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, userId: userId, userName: userName, calendarId: calendarId });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither userId nor userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, calendarId: calendarId });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, userId: 'foo', calendarId: calendarId });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, userName: 'foo', calendarId: calendarId });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both calendarId and calendarName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, userId: userId, calendarId: calendarId, calendarName: calendarName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: id, unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('retrieves event by id for the user specified with userId and calendarId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}/events/${id}`) {
        return eventResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: id,
        userId: userId,
        calendarId: calendarId,
        verbose: true
      })
    });
    assert(loggerLogSpy.calledOnceWith(eventResponse));
  });

  it('retrieves event by id for the user specified with userName and calendarName', async () => {
    sinon.stub(calendar, 'getUserCalendarByName').resolves({ id: calendarId });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${formatting.encodeQueryParameter(userName)}')/calendars/${calendarId}/events/${id}`) {
        return eventResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: id,
        userName: userName,
        calendarName: calendarName,
        timeZone: 'Pacific Standard Time',
        verbose: true
      })
    });
    assert(loggerLogSpy.calledOnceWith(eventResponse));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: id, userId: userId, calendarId: calendarId }) }),
      new CommandError(errorMessage)
    );
  });
});