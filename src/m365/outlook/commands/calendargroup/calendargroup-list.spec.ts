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
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './calendargroup-list.js';

describe(commands.CALENDARGROUP_LIST, () => {
  const userId = 'b743445a-112c-4fda-9afd-05943f9c7b36';
  const userName = 'john.doe@contoso.com';
  const currentUserId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
  const currentUserName = 'current.user@contoso.com';

  const calendarGroupsResponse = {
    value: [
      {
        "id": "AAMkAGE0MGM1Y2M5LWEzMmUtNGVlNy05MjRlLTk0YmYyY2I5NTM3ZAAuAAAAAAC_0WfqSjt_SqLtNkuO-bj1AQAbfYq5lmBxQ6a4t1fGbeYAAAAAAEOAAA=",
        "name": "My Calendars",
        "changeKey": "nfZyf7VcrEKLNoU37KWlkQAAA0x0+w==",
        "classId": "0006f0b7-0000-0000-c000-000000000046"
      },
      {
        "id": "AAMkAGE0MGM1Y2M5LWEzMmUtNGVlNy05MjRlLTk0YmYyY2I5NTM3ZAAuAAAAAAC_0WfqSjt_SqLtNkuO-bj1AQAbfYq5lmBxQ6a4t1fGbeYAAAAAAEPAAA=",
        "name": "Other Calendars",
        "changeKey": "nfZyf7VcrEKLNoU37KWlkQAAA0x0/A==",
        "classId": "0006f0b7-0000-0000-c000-000000000046"
      }
    ]
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDARGROUP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name']);
  });

  it('passes validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
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

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ userId: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ userName: 'foo' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('retrieves calendar groups for the signed-in user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/calendarGroups') {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('retrieves calendar groups for the signed-in user (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/calendarGroups') {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('retrieves calendar groups for a user specified by id using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('retrieves calendar groups for a user specified by user principal name using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userName }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('retrieves calendar groups for a user specified by id using app-only permissions (verbose)', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId, verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('retrieves calendar groups for a user specified by id using delegated permissions with Calendars.Read.Shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.Read.Shared']);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('retrieves calendar groups for a user specified by user principal name using delegated permissions with Calendars.Read.Shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.Read.Shared']);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userName }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('retrieves calendar groups for a user specified by id using delegated permissions with Calendars.ReadWrite.Shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('retrieves calendar groups for a user specified by id using delegated permissions with Calendars.Read.Shared scope (verbose)', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.Read.Shared']);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId, verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('does not check shared scope when userId matches the signed-in user', async () => {
    sinonUtil.restore(accessToken.getUserIdFromAccessToken);
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(userId);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('does not check shared scope when userName matches the signed-in user', async () => {
    sinonUtil.restore(accessToken.getUserNameFromAccessToken);
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns(userName);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups`) {
        return calendarGroupsResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ userName }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponse.value));
  });

  it('throws error when running with app-only permissions without userId or userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('When running with application permissions either userId or userName is required.')
    );
  });

  it('throws error when using delegated permissions with other userId without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ userId }) }),
      new CommandError(`To retrieve calendar groups of other users, the Entra ID application used for authentication must have either the Calendars.Read.Shared or Calendars.ReadWrite.Shared delegated permission assigned.`)
    );
  });

  it('throws error when using delegated permissions with other userName without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ userName }) }),
      new CommandError(`To retrieve calendar groups of other users, the Entra ID application used for authentication must have either the Calendars.Read.Shared or Calendars.ReadWrite.Shared delegated permission assigned.`)
    );
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError(errorMessage)
    );
  });
});
