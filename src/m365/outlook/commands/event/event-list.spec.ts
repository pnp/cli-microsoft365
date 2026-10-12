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
import command, { options } from './event-list.js';
import { calendar } from '../../../../utils/calendar.js';

describe(commands.EVENT_LIST, () => {
  const userId = 'b743445a-112c-4fda-9afd-05943f9c7b36';
  const userName = 'john.doe@contoso.com';
  const calendarId = 'AAMkAGI2AAATZQAAA=';
  const calendarName = 'My Calendar';

  const eventsResponse = [
    {
      "id": "AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAADSG3wPE27kUeySjmT5eRT8QcAfJKVL07sbkmIfHqjbDnRgQAAAgENAAAAfJKVL07sbkmIfHqjbDnRgQAC6GQ5pgAAAA==",
      "createdDateTime": "2026-03-29T13:57:47.9194633Z",
      "lastModifiedDateTime": "2026-03-29T13:59:48.6329479Z",
      "changeKey": "fJKVL07sbkmIfHqjbDnRgQAC54IeWA==",
      "categories": [],
      "transactionId": "localevent:c95ac848-7295-ad3e-ee1e-f3832b10bf3e",
      "originalStartTimeZone": "Greenwich Standard Time",
      "originalEndTimeZone": "Greenwich Standard Time",
      "iCalUId": "040000008200E00074C5B7101A82E008000000006B71750684BFDC01000000000000000010000000872F2916501A8442A7DB64D2E460E3D9",
      "uid": "040000008200E00074C5B7101A82E008000000006B71750684BFDC01000000000000000010000000872F2916501A8442A7DB64D2E460E3D9",
      "reminderMinutesBeforeStart": 15,
      "isReminderOn": true,
      "hasAttachments": false,
      "subject": "Pub",
      "bodyPreview": "sdfsdfsdfsdfdsfsdfsdfsd",
      "importance": "normal",
      "sensitivity": "normal",
      "isAllDay": false,
      "isCancelled": false,
      "isOrganizer": true,
      "responseRequested": true,
      "seriesMasterId": null,
      "showAs": "busy",
      "type": "singleInstance",
      "webLink": "https://outlook.office365.com/owa/?itemid=AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAADSG3wPE27kUeySjmT5eRT8QcAfJKVL07sbkmIfHqjbDnRgQAAAgENAAAAfJKVL07sbkmIfHqjbDnRgQAC6GQ5pgAAAA%3D%3D&exvsurl=1&path=/calendar/item",
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
        "content": "<html>\r\\\n<head>\r\\\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\\\n</head>\r\\\n<body>\r\\\n<div class=\"elementToProof\" style=\"font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)\">\r\\\nsdfsdfsdfsdfdsfsdfsdfsd</div>\r\\\n</body>\r\\\n</html>\r\\\n"
      },
      "start": {
        "dateTime": "2026-03-29T16:00:00.0000000",
        "timeZone": "UTC"
      },
      "end": {
        "dateTime": "2026-03-29T18:00:00.0000000",
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
          "name": "Martin Macháček",
          "address": "MartinMachacek@4wrvkx.onmicrosoft.com"
        }
      },
      "onlineMeeting": null
    },
    {
      "id": "AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAADSG3wPE27kUeySjmT5eRT8QcAfJKVL07sbkmIfHqjbDnRgQAAAgENAAAAfJKVL07sbkmIfHqjbDnRgQAC6GQ5pQAAAA==",
      "createdDateTime": "2026-03-29T13:57:18.3565941Z",
      "lastModifiedDateTime": "2026-03-29T13:57:19.5423408Z",
      "changeKey": "fJKVL07sbkmIfHqjbDnRgQAC54IdjA==",
      "categories": [],
      "transactionId": "localevent:0209423d-9958-b2db-5fcc-39360518b2b8",
      "originalStartTimeZone": "Greenwich Standard Time",
      "originalEndTimeZone": "Greenwich Standard Time",
      "iCalUId": "040000008200E00074C5B7101A82E00800000000C699D6F483BFDC0100000000000000001000000035C1CD3344304A40ACDF54500FE2F871",
      "uid": "040000008200E00074C5B7101A82E00800000000C699D6F483BFDC0100000000000000001000000035C1CD3344304A40ACDF54500FE2F871",
      "reminderMinutesBeforeStart": 15,
      "isReminderOn": true,
      "hasAttachments": false,
      "subject": "Testik",
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
      "webLink": "https://outlook.office365.com/owa/?itemid=AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAADSG3wPE27kUeySjmT5eRT8QcAfJKVL07sbkmIfHqjbDnRgQAAAgENAAAAfJKVL07sbkmIfHqjbDnRgQAC6GQ5pQAAAA%3D%3D&exvsurl=1&path=/calendar/item",
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
        "content": ""
      },
      "start": {
        "dateTime": "2026-03-30T14:30:00.0000000",
        "timeZone": "UTC"
      },
      "end": {
        "dateTime": "2026-03-30T15:00:00.0000000",
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
          "name": "Martin Macháček",
          "address": "MartinMachacek@4wrvkx.onmicrosoft.com"
        }
      },
      "onlineMeeting": null
    }
  ];

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
    assert.strictEqual(command.name, commands.EVENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'subject']);
  });

  it('passes validation with userId', () => {
    const actual = commandOptionsSchema.safeParse({ userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with userName', () => {
    const actual = commandOptionsSchema.safeParse({ userName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ userId, userName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither userId nor userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: calendarId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ userId: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid GUID.`);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ userName: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid UPN.`);
  });

  it('fails validation if both calendarId and calendarName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ calendarId, calendarName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if startDateTime is not a valid ISO date-time', () => {
    const actual = commandOptionsSchema.safeParse({ startDateTime: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid ISO date-time.`);
  });

  it('fails validation if endDateTime is not a valid ISO date-time', () => {
    const actual = commandOptionsSchema.safeParse({ endDateTime: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid ISO date-time.`);
  });

  it('retrieves events for the user specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/events`) {
        return {
          value: eventsResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId: userId, verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(eventsResponse));
  });

  it('retrieves filtered events in specific time zone for the user specified by UPN from a calendar specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendars/${calendarId}/events?$filter=contains(subject, 'contoso')`) {
        return {
          value: eventsResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        userName: userName,
        calendarId: calendarId,
        timeZone: 'Pacific Standard Time',
        filter: "contains(subject, 'contoso')",
        verbose: true
      })
    });
    assert(loggerLogSpy.calledOnceWith(eventsResponse));
  });

  it('retrieves filtered events since the specified date for the user specified by UPN from a calendar specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendars/${calendarId}/events?$filter=contains(subject, 'contoso') and start/dateTime ge '2026-03-29T00:00:00Z'`) {
        return {
          value: eventsResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        userName: userName,
        calendarId: calendarId,
        startDateTime: '2026-03-29T00:00:00Z',
        filter: "contains(subject, 'contoso')",
        verbose: true
      })
    });
    assert(loggerLogSpy.calledOnceWith(eventsResponse));
  });

  it('retrieves limited properties of events since the specified date for the user specified by id from a calendar specified by name', async () => {
    sinon.stub(calendar, 'getUserCalendarByName').resolves({ id: calendarId });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}/events?$select=id,subject,start,end&$filter=start/dateTime ge '2026-03-29T00:00:00Z'`) {
        return {
          value: eventsResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        userId: userId,
        calendarName: calendarName,
        startDateTime: '2026-03-29T00:00:00Z',
        properties: 'id,subject,start,end',
        verbose: true
      })
    });
    assert(loggerLogSpy.calledOnceWith(eventsResponse));
  });

  it('retrieves limited properties of events till the specified date for the user specified by id from a calendar specified by name', async () => {
    sinon.stub(calendar, 'getUserCalendarByName').resolves({ id: calendarId });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}/events?$select=id,subject,start,end&$filter=start/dateTime lt '2026-03-31T00:00:00Z'`) {
        return {
          value: eventsResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        userId: userId,
        calendarName: calendarName,
        endDateTime: '2026-03-31T00:00:00Z',
        properties: 'id,subject,start,end',
        verbose: true
      })
    });
    assert(loggerLogSpy.calledOnceWith(eventsResponse));
  });

  it('retrieves events in specific date range for the user specified by id', async () => {
    sinon.stub(calendar, 'getUserCalendarById').resolves({ id: calendarId });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/events?$expand=attachments($select=id)&$filter=start/dateTime ge '2026-03-29T00:00:00Z' and start/dateTime lt '2026-03-31T00:00:00Z'`) {
        return {
          value: eventsResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        userId: userId,
        startDateTime: '2026-03-29T00:00:00Z',
        endDateTime: '2026-03-31T00:00:00Z',
        properties: 'attachments/id',
        verbose: true
      })
    });
    assert(loggerLogSpy.calledOnceWith(eventsResponse));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ userId: userId }) }),
      new CommandError(errorMessage)
    );
  });
});