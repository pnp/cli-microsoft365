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
import command, { options } from './calendar-get.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.CALENDAR_GET, () => {
  const userId = 'ae0e8388-cd70-427f-9503-c57498ee3337';
  const userName = 'john.doe@contoso.com';
  const calendarId = 'AAMkADJmMVAAA=';
  const calendarName = 'Volunteer';
  const calendarGroupId = 'AQMkADJmMVAAA=';
  const calendarGroupName = 'My Calendars';
  const response = {
    "id": "AAMkAGI2TGuLAAA=",
    "name": "Calendar",
    "color": "auto",
    "isDefaultCalendar": true,
    "changeKey": "nfZyf7VcrEKLNoU37KWlkQAAA0x0+w==",
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDAR_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if neither id nor name is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both id and name is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: calendarId,
      name: calendarName,
      userId: userId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: calendarId,
      userId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid user principal name', () => {
    const actual = commandOptionsSchema.safeParse({
      id: calendarId,
      userName: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: calendarId,
      userId: userId,
      userName: userName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both calendarGroupId and calendarGroupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: calendarId,
      userId: userId,
      calendarGroupId: calendarGroupId,
      calendarGroupName: calendarGroupName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly retrieves a calendar by id for a user specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      id: calendarId,
      userId: userId,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly retrieves a calendar by id for a user specified by name from a calendar group specified by name', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').withArgs(userName, calendarGroupName, 'id').resolves({ id: calendarGroupId });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups/${calendarGroupId}/calendars/${calendarId}`) {
        return response;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      id: calendarId,
      userName: userName,
      calendarGroupName: calendarGroupName,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('correctly retrieves a calendar by name for a user specified by name from a calendar group specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups/${calendarGroupId}/calendars?$filter=name eq '${formatting.encodeQueryParameter(calendarName)}'`) {
        return {
          value: [response]
        };
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      name: calendarName,
      userName: userName,
      calendarGroupId: calendarGroupId,
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('handles error when calendar was not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}`) {
        throw {
          error:
          {
            code: 'Request_ResourceNotFound',
            message: `Resource '${calendarId}' does not exist or one of its queried reference-property objects are not present.`
          }
        };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: { id: calendarId, userId: userId } }),
      new CommandError(`Resource '${calendarId}' does not exist or one of its queried reference-property objects are not present.`)
    );
  });
});
