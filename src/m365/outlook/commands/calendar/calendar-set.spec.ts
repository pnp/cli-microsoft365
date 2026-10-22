import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './calendar-set.js';

describe(commands.CALENDAR_SET, () => {
  const calendarId = 'AAMkAGI2TQpZAAA=';
  const userId = 'b743445a-112c-4fda-9afd-05943f9c7b36';
  const userName = 'john.doe@contoso.com';
  const currentUserId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
  const currentUserName = 'current.user@contoso.com';
  const calendarGroupId = 'AAMkAGI2TGuMAAA=';
  const calendarGroupName = 'My Calendars';

  const calendarResponse = {
    "id": calendarId,
    "name": "Team planning",
    "color": "auto",
    "hexColor": "",
    "isDefaultCalendar": false,
    "changeKey": "DxYSthXJXEWwAQSYQnXvIgAAIxGttg==",
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns([]);
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(currentUserId);
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns(currentUserName);
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      accessToken.getScopesFromAccessToken,
      accessToken.getUserIdFromAccessToken,
      accessToken.getUserNameFromAccessToken,
      request.patch,
      calendarGroup.getUserCalendarGroupByName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDAR_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation with id and name', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id and color', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, color: 'lightGreen' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id and isDefault', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, isDefault: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id, name and userId', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id, name and userName', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', userName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id, name and calendarGroupId', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', calendarGroupId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id, name and calendarGroupName', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', calendarGroupName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if id is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'Team planning' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if no updatable property is provided', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', userId, userName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both calendarGroupId and calendarGroupName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', calendarGroupId, calendarGroupName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', userId: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid GUID.`);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', userName: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid UPN.`);
  });

  it('fails validation if color has an invalid value', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, color: 'invalidColor' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarId, name: 'Team planning', unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('updates the name of a calendar for the signed-in user', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendars/${calendarId}` && JSON.stringify(opts.data) === JSON.stringify({ name: 'Team planning' })) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning' }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates the color of a calendar for the signed-in user', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendars/${calendarId}` && JSON.stringify(opts.data) === JSON.stringify({ color: 'lightGreen' })) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, color: 'lightGreen' }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('marks a calendar as default for the signed-in user', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendars/${calendarId}` && JSON.stringify(opts.data) === JSON.stringify({ isDefaultCalendar: true })) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, isDefault: true }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates a calendar for the signed-in user (verbose)', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates a calendar for a user specified by id using delegated permissions with shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates a calendar for a user specified by UPN using delegated permissions with shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', userName }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('does not check shared scope when userId matches the signed-in user', async () => {
    sinonUtil.restore(accessToken.getUserIdFromAccessToken);
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(userId);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('does not check shared scope when userName matches the signed-in user', async () => {
    sinonUtil.restore(accessToken.getUserNameFromAccessToken);
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns(userName);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', userName }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates a calendar for a user specified by id using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates a calendar for a user specified by UPN using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', userName }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates a calendar within a specific calendar group by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendarGroups/${calendarGroupId}/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', calendarGroupId }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates a calendar within a specific calendar group by name', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').resolves({ id: calendarGroupId });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendarGroups/${calendarGroupId}/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', calendarGroupName }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('updates a calendar within a specific calendar group by name for a specific user', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').resolves({ id: calendarGroupId });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}/calendars/${calendarId}`) {
        return calendarResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', calendarGroupName, userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarResponse));
  });

  it('throws error when running with app-only permissions without userId or userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning' }) }),
      new CommandError('When running with application permissions either userId or userName is required.')
    );
  });

  it('throws error when using delegated permissions with other userId without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', userId }) }),
      new CommandError('To update calendars of other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.')
    );
  });

  it('throws error when using delegated permissions with other userName without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', userName }) }),
      new CommandError('To update calendars of other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.')
    );
  });

  it('throws error when calendar group name does not exist', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').rejects(new Error(`The specified calendar group '${calendarGroupName}' does not exist.`));

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning', calendarGroupName }) }),
      new CommandError(`The specified calendar group '${calendarGroupName}' does not exist.`)
    );
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'patch').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarId, name: 'Team planning' }) }),
      new CommandError(errorMessage)
    );
  });
});
