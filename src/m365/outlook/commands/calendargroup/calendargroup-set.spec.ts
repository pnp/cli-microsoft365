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
import command, { options } from './calendargroup-set.js';
import { z } from 'zod';

describe(commands.CALENDARGROUP_SET, () => {
  const calendarGroupId = 'AAMkAGE0MGM1Y2M5LWEzMmUtNGVlNy05MjRlLTk0YmYyY2I5NTM3ZAAuAAAAAAC_0WfqSjt_SqLtNkuO-bj1AQAbfYq5lmBxQ6a4t1fGbeYAAAAAAEOAAA=';
  const calendarGroupName = 'My Calendars';
  const newName = 'Personal Events';
  const userId = 'b743445a-112c-4fda-9afd-05943f9c7b36';
  const userName = 'john.doe@contoso.com';
  const currentUserId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
  const currentUserName = 'current.user@contoso.com';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let refinedSchema: z.ZodTypeAny;

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
    refinedSchema = commandInfo.command.getRefinedSchema!(commandOptionsSchema as any)!;
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
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDARGROUP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation with id and newName', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, newName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with name and newName', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName, newName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id, newName and userId', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, newName, userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id, newName and userName', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, newName, userName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = refinedSchema.safeParse({ id: calendarGroupId, newName, userId, userName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither id nor name is specified', () => {
    const actual = refinedSchema.safeParse({ newName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both id and name are specified', () => {
    const actual = refinedSchema.safeParse({ id: calendarGroupId, name: calendarGroupName, newName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, newName, userId: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid GUID.`);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, newName, userName: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid UPN.`);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, newName, unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('updates a calendar group by id for the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('updates a calendar group by name for the signed-in user', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').resolves({ id: calendarGroupId, name: calendarGroupName });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, newName }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('updates a calendar group by id for the signed-in user (verbose)', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, verbose: true }) });
    assert(log.some((l: any) => l.indexOf('Updating calendar group for the signed-in user...') > -1));
    assert(log.some((l: any) => l.indexOf(`Updating calendar group '${calendarGroupId}'...`) > -1));
  });

  it('updates a calendar group by id for a user specified by userId using app-only permissions (verbose)', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userId, verbose: true }) });
    assert(log.some((l: any) => l.indexOf(`Updating calendar group using application permissions for user '${userId}'...`) > -1));
  });

  it('updates a calendar group by id for a user specified by userId using delegated permissions (verbose)', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userId, verbose: true }) });
    assert(log.some((l: any) => l.indexOf(`Updating calendar group using delegated permissions for user '${userId}'...`) > -1));
  });

  it('updates a calendar group by id for a user specified by userId using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userId }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('updates a calendar group by id for a user specified by userName using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('john.doe%40contoso.com')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userName }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('updates a calendar group by name for a user specified by userId using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').resolves({ id: calendarGroupId, name: calendarGroupName });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, newName, userId }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('updates a calendar group by id for a user specified by userId using delegated permissions with Calendars.ReadWrite.Shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userId }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('updates a calendar group by id for a user specified by userName using delegated permissions with Calendars.ReadWrite.Shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('john.doe%40contoso.com')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userName }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('does not check shared scope when userId matches the signed-in user', async () => {
    sinonUtil.restore(accessToken.getUserIdFromAccessToken);
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(userId);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userId }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('does not check shared scope when userName matches the signed-in user', async () => {
    sinonUtil.restore(accessToken.getUserNameFromAccessToken);
    sinon.stub(accessToken, 'getUserNameFromAccessToken').returns(userName);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('john.doe%40contoso.com')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userName }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { name: newName });
  });

  it('throws error when running with app-only permissions without userId or userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName }) }),
      new CommandError('When running with application permissions either userId or userName is required.')
    );
  });

  it('throws error when using delegated permissions with other userId without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userId }) }),
      new CommandError('To update calendar groups of other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.')
    );
  });

  it('throws error when using delegated permissions with other userName without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName, userName }) }),
      new CommandError('To update calendar groups of other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.')
    );
  });

  it('throws error when calendar group with specified name is not found', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').rejects(new Error("The specified calendar group 'NonExistent Group' does not exist."));

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ name: 'NonExistent Group', newName }) }),
      new CommandError("The specified calendar group 'NonExistent Group' does not exist.")
    );
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'patch').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, newName }) }),
      new CommandError(errorMessage)
    );
  });
});
