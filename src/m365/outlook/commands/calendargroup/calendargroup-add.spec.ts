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
import command, { options } from './calendargroup-add.js';

describe(commands.CALENDARGROUP_ADD, () => {
  const calendarGroupName = 'My Work Calendars';
  const userId = 'b743445a-112c-4fda-9afd-05943f9c7b36';
  const userName = 'john.doe@contoso.com';
  const currentUserId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
  const currentUserName = 'current.user@contoso.com';
  const response = {
    "id": "AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAADSG3wPE27kUeySjmT5eRT8QcAfJKVL07sbkmIfHqjbDnRgQAAAgEGAAAAfJKVL07sbkmIfHqjbDnRgQADK5c4ngAAAA==",
    "name": "My Work Calendars",
    "classId": "c02d4ddf-4809-485f-9cd4-1ef0937e03be",
    "changeKey": "fJKVL07sbkmIfHqjbDnRgQADKqwfAA=="
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
      calendarGroup.getUserCalendarGroupByName,
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDARGROUP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation with name and userId', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName, userId: userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with name and userName', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName, userName: userName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName, userId: userId, userName: userName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if name is specified', () => {
    const actual = commandOptionsSchema.safeParse({ userId: userId });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName, userId: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName, userName: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName, unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('creates a calendar group for the signed-in user', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/calendarGroups') {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('creates a calendar group for the signed-in user (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/calendarGroups') {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, verbose: true }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('creates a calendar group for a user specified by userId using app-only permissions (verbose)', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userId: userId, verbose: true }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('creates a calendar group for a user specified by userId using delegated permissions (verbose)', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userId: userId, verbose: true }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('creates a calendar group for a user specified by userId using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userId: userId }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('creates a calendar group for a user specified by userName using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('john.doe%40contoso.com')/calendarGroups`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userName: userName }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('creates a calendar group for a user specified by userId using delegated permissions with Calendars.ReadWrite.Shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userId: userId }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('creates a calendar group for a user specified by userName using delegated permissions with Calendars.ReadWrite.Shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('john.doe%40contoso.com')/calendarGroups`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userName: userName }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('does not check shared scope when userId matches the signed-in user', async () => {
    sinonUtil.restore(accessToken.getUserIdFromAccessToken);
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(userId);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userId: userId }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('does not check shared scope when userName matches the signed-in user', async () => {
    sinonUtil.restore(accessToken.getUserNameFromAccessToken);
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns(userName);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('john.doe%40contoso.com')/calendarGroups`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userName: userName }) });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('throws error when running with app-only permissions without userId or userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName }) }),
      new CommandError('When running with application permissions either userId or userName is required.')
    );
  });

  it('throws error when using delegated permissions with other userId without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userId: userId }) }),
      new CommandError('To create calendar groups of other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.')
    );
  });

  it('throws error when using delegated permissions with other userName without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userName: userName }) }),
      new CommandError('To create calendar groups of other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.')
    );
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'post').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName }) }),
      new CommandError(errorMessage)
    );
  });
});
