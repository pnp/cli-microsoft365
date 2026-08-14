import assert from 'assert';
import fs from 'fs';
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
import command, { options } from './event-add.js';
import { calendar } from '../../../../utils/calendar.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.EVENT_ADD, () => {
  const subject = 'CLI sync';
  const start = '2026-08-06T10:00:00';
  const end = '2026-08-06T11:00:00';
  const userId = "9bd29c6c-181e-41f5-a1b6-bc30bbf652d3";
  const userName = "john.doe@contoso.com";
  const calendarId = "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAAAAAEGAADEwEFouXWWT50CfwqSN9cpAAAkuACjAAA=";
  const calendarName = "Calendar";
  const response = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('9bd29c6c-181e-41f5-a1b6-bc30bbf652d3')/events/$entity",
    "@odata.etag": "W/\"AgHex1lHnU6P4tmSFYQ4hwAL1DoJsQ==\"",
    "id": "AAMkAGRkZTFiMDQxLWYzNDgtNGQ2U3LWU1NWJhMTM5YTgwMABGAAAAAABxI4iNfZK7SYRiWw9s0BwA7DGC6yx9ARZqQFWs3P3q1AAAOQAAACAd7HWUedTo-i2ZIVhDAAvXjiwwAAA=",
    "createdDateTime": "2026-08-06T10:04:43.8235629Z",
    "lastModifiedDateTime": "2026-08-06T10:04:43.9121602Z",
    "changeKey": "AgHex1l6P4tmSFYQ4hwAL1DoJsQ==",
    "categories": [],
    "transactionId": null,
    "originalStartTimeZone": "UTC",
    "originalEndTimeZone": "UTC",
    "iCalUId": "040000008200E00074C5B7101A82E008000000005F63B25DD01000000000000000010000000D3F5CCB8E212D441A51969663E1CAC1A",
    "uid": "040000008200E00074C5B7101A82E008000000005F63B25DD01000000000000000010000000D3F5CCB8E212D441A51969663E1CAC1A",
    "reminderMinutesBeforeStart": 15,
    "isReminderOn": true,
    "hasAttachments": false,
    "subject": "Sync",
    "bodyPreview": "",
    "importance": "normal",
    "sensitivity": "normal",
    "isAllDay": false,
    "isCancelled": false,
    "isOrganizer": true,
    "responseRequested": true,
    "seriesMasterId": null,
    "showAs": "busy",
    "type": "singleInstance",
    "webLink": "https://outlook.office365.com/owa/?itemid=AAMkAGRAAA%3D&exvsurl=1&path=/calendar/item",
    "onlineMeetingUrl": null,
    "isOnlineMeeting": false,
    "onlineMeetingProvider": "unknown",
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
      "content": "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\n<meta name=\"Generator\" content=\"Microsoft Exchange Server\">\r\n<!-- converted from text -->\r\n<style><!-- .EmailQuote { margin-left: 1pt; padding-left: 4pt; border-left: #800000 2px solid; } --></style></head>\r\n<body>\r\n<font size=\"2\"><span style=\"font-size:11pt;\"><div class=\"PlainText\">&nbsp;</div></span></font>\r\n</body>\r\n</html>\r\n"
    },
    "start": {
      "dateTime": "2026-08-06T10:00:00.0000000",
      "timeZone": "UTC"
    },
    "end": {
      "dateTime": "2026-08-06T10:15:00.0000000",
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
    "attendees": [],
    "organizer": {
      "emailAddress": {
        "name": "John doe",
        "address": "john.doe@contoso.com"
      }
    },
    "onlineMeeting": null
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      calendar.getUserCalendarByName,
      accessToken.isAppOnlyAccessToken,
      accessToken.getUserIdFromAccessToken,
      accessToken.getUserNameFromAccessToken,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EVENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation with text body content type', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, bodyContentType: 'Text' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with html body content type', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, bodyContentType: 'HTML' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with low importance', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, importance: 'low' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with normal importance', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, importance: 'normal' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with high importance', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, importance: 'high' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with normal sensitivity', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, sensitivity: 'normal' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with personal sensitivity', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, sensitivity: 'personal' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with private sensitivity', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, sensitivity: 'private' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with confidential sensitivity', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, sensitivity: 'confidential' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with teamsForBusiness online meeting provider', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, onlineMeetingProvider: 'teamsForBusiness' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with skypeForBusiness online meeting provider', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, onlineMeetingProvider: 'skypeForBusiness' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with skypeForConsumer online meeting provider', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, onlineMeetingProvider: 'skypeForConsumer' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with unknown online meeting provider', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, onlineMeetingProvider: 'unknown' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for free status', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, showAs: 'free' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for tentative status', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, showAs: 'tentative' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for busy status', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, showAs: 'busy' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for out of office status', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, showAs: 'oof' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for working elsewhere status', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, showAs: 'workingElsewhere' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for unknown status', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, showAs: 'unknown' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if user id is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, userId: userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if user name is a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, userName: userName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if reminderMinutesBeforeStart is 0', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, reminderMinutesBeforeStart: 0 });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if reminderMinutesBeforeStart is greater than 0', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, reminderMinutesBeforeStart: 15, isReminderOn: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if optionalAttendees contains valid email addresses', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, optionalAttendees: 'john.doe@contoso.com,alice.weber@constoso.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if requiredAttendees contains valid email addresses', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, requiredAttendees: 'john.doe@contoso.com,alice.weber@constoso.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if resources contains valid email addresses', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, resources: 'room1@contoso.com,room2@constoso.com' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if start date time is not a valid ISO 8601 date', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: 'foo', end: end });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if end date time is not a valid ISO 8601 date', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation for incorrect body content type', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, bodyContentType: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation for incorrect importance', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, importance: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation for incorrect sensitivity', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, sensitivity: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation for incorrect online meeting provider', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, onlineMeetingProvider: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation for incorrect status', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, showAs: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if user id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, userId: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if user name is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, userName: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, userId: userId, userName: userName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both calendarId and calendarName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, calendarId: calendarId, calendarName: calendarName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both location and locations are specified', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, location: 'Meeting Room 1', locations: 'Meeting Room 1,Meeting Room 2' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if locationEmailAddress is specified, but location is missing', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, locationEmailAddress: 'meetingRoom1@contoso.com' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if locationEmailAddress is not a valid email address', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, locationEmailAddress: 'foo', location: 'Meeting Room 1' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if reminderMinutesBeforeStart is specified, but isReminderOn is false', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, reminderMinutesBeforeStart: 5, isReminderOn: false });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if isAllDay is true, but start is not set to midnight', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: '2026-08-08T00:00:00', isAllDay: true });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if isAllDay is true, but end is not set to midnight', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: '2026-08-05T00:00:00', end: end, isAllDay: true });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if end starts before start', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: '2026-01-01T00:00:00' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if reminderMinutesBeforeStart is negative', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, reminderMinutesBeforeStart: -5 });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if optionalAttendees contains invalid email address', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, optionalAttendees: 'john.doe@contoso.com,foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if requiredAttendees contains invalid email address', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, requiredAttendees: 'john.doe@contoso.com,foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if resources contains invalid email address', () => {
    const actual = commandOptionsSchema.safeParse({ subject: subject, start: start, end: end, resources: 'john.doe@contoso.com,foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly creates an event if only subject, start and end are specified', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event for a user specified by id in a calendar specified by id', async () => {
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(userId);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      userId: userId,
      calendarId: calendarId,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event for a user specified by UPN in a calendar specified by name', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns(userName);
    sinon.stub(calendar, 'getUserCalendarByName').resolves({ id: calendarId }).calledWith(calendarName);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendars/${calendarId}/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      userName: userName,
      calendarName: calendarName,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with optional and required attendees and with resources', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      optionalAttendees: 'user1@contoso.com,user2@contoso.com',
      requiredAttendees: 'user3@contoso.com, user4@contoso.com',
      resources: 'meetingRoom1@contoso.com,meetingRoom2@contoso.com',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      attendees: [
        {
          emailAddress: {
            address: 'user1@contoso.com'
          },
          type: 'optional'
        },
        {
          emailAddress: {
            address: 'user2@contoso.com'
          },
          type: 'optional'
        },
        {
          emailAddress: {
            address: 'user3@contoso.com'
          },
          type: 'required'
        },
        {
          emailAddress: {
            address: 'user4@contoso.com'
          },
          type: 'required'
        },
        {
          emailAddress: {
            address: 'meetingRoom1@contoso.com'
          },
          type: 'resource'
        },
        {
          emailAddress: {
            address: 'meetingRoom2@contoso.com'
          },
          type: 'resource'
        }
      ]
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with attendees and hide them', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      requiredAttendees: 'user1@contoso.com,user2@contoso.com',
      hideAttendees: true,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      attendees: [
        {
          emailAddress: {
            address: 'user1@contoso.com'
          },
          type: 'required'
        },
        {
          emailAddress: {
            address: 'user2@contoso.com'
          },
          type: 'required'
        }
      ],
      hideAttendees: true
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with allowNewTimeProposals, isReminderOn, responseRequested, categories, sensitivity, importance, showAs and transactionId', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      allowNewTimeProposals: false,
      isReminderOn: false,
      responseRequested: false,
      categories: 'category1,category2',
      sensitivity: 'private',
      importance: 'high',
      showAs: 'oof',
      transactionId: 'xxx',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      allowNewTimeProposals: false,
      isReminderOn: false,
      responseRequested: false,
      categories: [
        'category1',
        'category2'
      ],
      sensitivity: 'private',
      importance: 'high',
      showAs: 'oof',
      transactionId: 'xxx'
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with location and locationEmailAddress', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      location: 'Room 1',
      locationEmailAddress: 'meetingRoom1@contoso.com',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      location: {
        displayName: 'Room 1',
        locationEmailAddress: 'meetingRoom1@contoso.com'
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with locations', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      locations: 'Room 1,Room 2',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      locations: [
        {
          displayName: 'Room 1'
        },
        {
          displayName: 'Room 2'
        }
      ]
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with all day online meeting', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: '2026-08-06T00:00:00',
      end: '2026-08-07T00:00:00',
      isOnlineMeeting: true,
      isAllDay: true,
      onlineMeetingProvider: 'teamsForBusiness',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T00:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-07T00:00:00',
        timeZone: 'UTC'
      },
      isOnlineMeeting: true,
      isAllDay: true,
      onlineMeetingProvider: 'teamsForBusiness'
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with a reminder', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      isReminderOn: true,
      reminderMinutesBeforeStart: 30,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      reminderMinutesBeforeStart: 30
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates a reccuring event', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      recurrence: `{ "pattern": { "type": "weekly", "interval": 1, "daysOfWeek": [ "Monday" ] }, "range": { "type": "endDate", "startDate": "2017-09-04", "endDate": "2017-12-31" } }`,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      recurrence: {
        "pattern": {
          "type": "weekly",
          "interval": 1,
          "daysOfWeek": ["Monday"]
        },
        "range": {
          "type": "endDate",
          "startDate": "2017-09-04",
          "endDate": "2017-12-31"
        }
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event and reads an occurence from a file', async () => {
    const reccurence = {
      "pattern": {
        "type": "weekly",
        "interval": 1,
        "daysOfWeek": [
          "Monday"
        ]
      }, "range": {
        "type": "endDate",
        "startDate": "2026-08-08",
        "endDate": "2026-08-24"
      }
    };
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify(reccurence));
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      recurrence: '@file',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      recurrence: {
        "pattern": {
          "type": "weekly",
          "interval": 1,
          "daysOfWeek": ["Monday"]
        },
        "range": {
          "type": "endDate",
          "startDate": "2026-08-08",
          "endDate": "2026-08-24"
        }
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with text body', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      bodyContents: `Let's go for lunch`,
      bodyContentType: 'Text',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      body: {
        content: `Let's go for lunch`,
        contentType: 'Text'
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates an event with html body from a file', async () => {
    const htmlText = `<html><h1>Let's go for lunch</h1></html>`;
    sinon.stub(fs, 'readFileSync').returns(htmlText);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      bodyContents: '@file',
      bodyContentType: 'HTML',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      subject: 'CLI sync',
      start: {
        dateTime: '2026-08-06T10:00:00',
        timeZone: 'UTC'
      },
      end: {
        dateTime: '2026-08-06T11:00:00',
        timeZone: 'UTC'
      },
      body: {
        content: `<html><h1>Let's go for lunch</h1></html>`,
        contentType: 'HTML'
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('throws an error when userId does not match current user when using delegated permissions', async () => {
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns('00000000-0000-0000-0000-000000000000');

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      userId: userId,
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }),
      new CommandError(`You can only create your own events when using delegated permissions. The specified userId '${userId}' does not match the current user '00000000-0000-0000-0000-000000000000'.`));
  });

  it('throws an error when userName does not match current user when using delegated permissions', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns('other.user@contoso.com');

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      userName: userName,
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }),
      new CommandError(`You can only create your own events when using delegated permissions. The specified userName '${userName}' does not match the current user 'other.user@contoso.com'.`));
  });

  it('throws an error when both userId and userName are not defined when creating an event using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }),
      new CommandError(`The option 'userId' or 'userName' is required when creating an event using application permissions.`));
  });

  it('correctly handles API errors', async () => {
    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `The specified object was not found in the store., The process failed to get the correct properties.`,
        innerError: {
          date: '2023-10-27T12:24:36',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendars/${calendarId}/events`) {
        throw error;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      subject: subject,
      start: start,
      end: end,
      calendarId: calendarId,
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }),
      new CommandError(error.error.message));
  });
});