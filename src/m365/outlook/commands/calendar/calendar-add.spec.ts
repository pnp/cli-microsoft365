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
import command, { options } from './calendar-add.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';

describe(commands.CALENDAR_ADD, () => {
  const userId = 'ae0e8388-cd70-427f-9503-c57498ee3337';
  const userName = 'john.doe@contoso.com';
  const calendarName = 'Volunteer';
  const calendarGroupId = 'AQMkADJmMVAAA=';
  const calendarGroupName = 'My Calendars';
  const response = {
    "id": "AAMkADJmMVAAA=",
    "name": "Volunteer",
    "color": "auto",
    "changeKey": "DxYSthXJXEWwAQSYQnXvIgAAIxGttg==",
    "canShare": true,
    "canViewPrivateItems": true,
    "hexColor": "",
    "canEdit": true,
    "allowedOnlineMeetingProviders": [
      "teamsForBusiness"
    ],
    "defaultOnlineMeetingProvider": "teamsForBusiness",
    "isTallyingResponses": true,
    "isRemovable": false,
    "owner": {
      "name": "John Doe",
      "address": "john.doe@contoso.com"
    }
  };

  let log: any[];
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDAR_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if name is not provided', () => {
    const actual = commandOptionsSchema.safeParse({ userId: userId });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: 'foo',
      name: calendarName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid user principal name', () => {
    const actual = commandOptionsSchema.safeParse({
      userName: 'foo',
      name: calendarName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      userName: userName,
      name: calendarName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if color is invalid', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      name: calendarName,
      color: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if defaultOnlineMeetingProvider is invalid', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      name: calendarName,
      defaultOnlineMeetingProvider: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly creates a calendar for a user specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      userId: userId,
      name: calendarName,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates a calendar for a user specified by UPN in a calendar group specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups/${calendarGroupId}/calendars`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      userName: userName,
      name: calendarName,
      calendarGroupId: calendarGroupId,
      defaultOnlineMeetingProvider: 'none',
      color: 'lightBlue'
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly creates a calendar for a user specified by UPN in a calendar group specified by name', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').withArgs(userName, calendarGroupName, 'id').resolves({ id: calendarGroupId });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups/${calendarGroupId}/calendars`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      userName: userName,
      name: calendarName,
      calendarGroupName: calendarGroupName
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      userId: userId,
      name: calendarName
    });
    await assert.rejects(command.action(logger, {
      options: parsedSchema.data!
    }), new CommandError('Invalid request'));
  });
});