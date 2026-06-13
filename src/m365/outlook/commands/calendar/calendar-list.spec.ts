import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './calendar-list.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';
import { odata } from '../../../../utils/odata.js';

describe(commands.CALENDAR_LIST, () => {
  const userId = 'ae0e8388-cd70-427f-9503-c57498ee3337';
  const userName = 'john.doe@contoso.com';
  const calendarGroupId = 'AQMkADJmMVAAA=';
  const calendarGroupName = 'My Calendars';
  const response = [
    {
      "id": "AAMkAGI2MDc2YzA0LWQwNTktNGM5Ni05M2VkLWY3NjFkNTUxOTkyZABGAAAAAABeGJMObKvfQbq5qwfGa7kTBwAopDdmUXY8TaLJk5CCLo4zAAAAAAEGAAAopDdmUXY8TaLJk5CCLo4zAAAAAFS0AAA=",
      "name": "Calendar",
      "color": "auto",
      "hexColor": "",
      "groupClassId": "0006f0b7-0000-0000-c000-000000000046",
      "isDefaultCalendar": true,
      "changeKey": "KKQ3ZlF2PE2iyZOQgi6OMwAAAAADcg==",
      "canShare": true,
      "canViewPrivateItems": true,
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
    },
    {
      "id": "AAMkAGI2MDc2YzA0LWQwNTktNGM5Ni05M2VkLWY3NjFkNTUxOTkyZABGAAAAAABeGJMObKvfQbq5qwfGa7kTBwAopDdmUXY8TaLJk5CCLo4zAAAAAAEGAAAopDdmUXY8TaLJk5CCLo4zAAAAAFS1AAA=",
      "name": "United States holidays",
      "color": "auto",
      "hexColor": "",
      "groupClassId": "0006f0b7-0000-0000-c000-000000000046",
      "isDefaultCalendar": false,
      "changeKey": "KKQ3ZlF2PE2iyZOQgi6OMwAAAAADfA==",
      "canShare": false,
      "canViewPrivateItems": true,
      "canEdit": false,
      "allowedOnlineMeetingProviders": [],
      "defaultOnlineMeetingProvider": "unknown",
      "isTallyingResponses": false,
      "isRemovable": true,
      "owner": {
        "name": "John Doe",
        "address": "john.doe@contoso.com"
      }
    },
    {
      "id": "AAMkAGI2MDc2YzA0LWQwNTktNGM5Ni05M2VkLWY3NjFkNTUxOTkyZABGAAAAAABeGJMObKvfQbq5qwfGa7kTBwAopDdmUXY8TaLJk5CCLo4zAAAAAAEGAAAopDdmUXY8TaLJk5CCLo4zAAAAAFS4AAA=",
      "name": "Birthdays",
      "color": "auto",
      "hexColor": "",
      "groupClassId": "0006f0b7-0000-0000-c000-000000000046",
      "isDefaultCalendar": false,
      "changeKey": "KKQ3ZlF2PE2iyZOQgi6OMwAAAAAFKg==",
      "canShare": false,
      "canViewPrivateItems": true,
      "canEdit": false,
      "allowedOnlineMeetingProviders": [],
      "defaultOnlineMeetingProvider": "unknown",
      "isTallyingResponses": false,
      "isRemovable": true,
      "owner": {
        "name": "John Doe",
        "address": "john.doe@contoso.com"
      }
    }
  ];

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
      odata.getAllItems,
      calendarGroup.getUserCalendarGroupByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDAR_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name']);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ userId, userName });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if only userId is specified', () => {
    const actual = commandOptionsSchema.safeParse({ userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if only userName is specified', () => {
    const actual = commandOptionsSchema.safeParse({ userName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ userId: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ userName: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both calendarGroupId and calendarGroupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({ userId, calendarGroupId, calendarGroupName });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if only calendarGroupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({ userId, calendarGroupId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if only calendarGroupName is specified', () => {
    const actual = commandOptionsSchema.safeParse({ userId, calendarGroupName });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves calendars for a user by userId', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId, verbose: true }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('retrieves calendars for a user by userName', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendars`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userName, verbose: true }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('retrieves calendars for a user and calendar group by calendarGroupId', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}/calendars`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId, calendarGroupId }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('retrieves calendars for a user and calendar group by calendarGroupName', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').withArgs(userName, calendarGroupName, 'id').resolves({ id: calendarGroupId });
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups/${calendarGroupId}/calendars`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userName, calendarGroupName }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('handles error when calendar was not found', async () => {
    const invalidUserName = 'invalidUser@contoso.com';
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users('${invalidUserName}')/calendars`) {
        throw {
          error:
          {
            code: 'ErrorInvalidUser',
            message: `The requested user '${invalidUserName}' is invalid.`
          }
        };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: { userName: invalidUserName } }),
      new CommandError(`The requested user '${invalidUserName}' is invalid.`)
    );
  });
});
