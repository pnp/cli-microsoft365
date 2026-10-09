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
import { formatting } from '../../../../utils/formatting.js';
import commands from '../../commands.js';
import command, { options } from './calendargroup-get.js';

describe(commands.CALENDARGROUP_GET, () => {
  const calendarGroupId = 'AAMkAGE0MGM1Y2M5LWEzMmUtNGVlNy05MjRlLTk0YmYyY2I5NTM3ZAAuAAAAAAC_0WfqSjt_SqLtNkuO-bj1AQAbfYq5lmBxQ6a4t1fGbeYAAAAAAEOAAA=';
  const calendarGroupName = 'Personal Events';
  const otherUserId = '44288f7d-7710-4293-8c8e-36f310ed2e6a';
  const userId = 'b743445a-112c-4fda-9afd-05943f9c7b36';
  const userName = 'john.doe@contoso.com';
  const currentUserId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
  const currentUserName = 'current.user@contoso.com';

  const calendarGroupResponse = {
    id: calendarGroupId,
    name: 'My Calendars',
    changeKey: 'nfZyf7VcrEKLNoU37KWlkQAAA0x0+w==',
    classId: '0006f0b7-0000-0000-c000-000000000046'
  };

  const calendarGroupsResponseForFilter = {
    value: [
      {
        id: calendarGroupId,
        name: calendarGroupName
      }
    ]
  };

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
    logger = {
      log: async () => undefined,
      logRaw: async () => undefined,
      logToStderr: async () => undefined
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
    assert.strictEqual(command.name, commands.CALENDARGROUP_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation with id', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with name', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both id and name are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, name: calendarGroupName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither id nor name is specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if id is empty', () => {
    const actual = commandOptionsSchema.safeParse({ id: '' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if name is empty', () => {
    const actual = commandOptionsSchema.safeParse({ name: '' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, userId: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid GUID.`);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, userName: 'foo' });
    assert.notStrictEqual(actual.success, true);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid UPN.`);
  });

  it('fails validation if both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, userId: userId, userName: userName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('retrieves calendar group for the signed-in user by id using delegated permissions', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendarGroups/${calendarGroupId}`) {
        return calendarGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupResponse));
  });

  it('retrieves calendar group for the signed-in user by name using delegated permissions', async () => {
    const expectedFilterUrl = `https://graph.microsoft.com/v1.0/me/calendarGroups?$filter=name eq 'Personal%20Events'`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === expectedFilterUrl) {
        return calendarGroupsResponseForFilter;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponseForFilter.value[0]));
  });

  it('retrieves calendar group for a user specified by id using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return calendarGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, userId }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupResponse));
  });

  it('throws error when running with app-only permissions without userId or userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId }) }),
      new CommandError('When running with application permissions either userId or userName is required.')
    );
  });

  it('throws error when using delegated permissions for other users without shared scope', async () => {
    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, userId: otherUserId }) }),
      new CommandError(`To retrieve calendar groups of other users, the Entra ID application used for authentication must have either the Calendars.Read.Shared or Calendars.ReadWrite.Shared delegated permission assigned.`)
    );
  });

  it('retrieves calendar group for a user specified by id using delegated permissions with shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.Read.Shared']);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${otherUserId}')/calendarGroups/${calendarGroupId}`) {
        return calendarGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, userId: otherUserId }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupResponse));
  });

  it('retrieves calendar group for a user specified by name using delegated permissions with shared scope', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.Read.Shared']);

    const expectedFilterUrl = `https://graph.microsoft.com/v1.0/users('${otherUserId}')/calendarGroups?$filter=name eq 'Personal%20Events'`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === expectedFilterUrl) {
        return calendarGroupsResponseForFilter;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userId: otherUserId }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponseForFilter.value[0]));
  });

  it('retrieves calendar group for the signed-in user with verbose output', async () => {
    const logToStderrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/calendarGroups/${calendarGroupId}`) {
        return calendarGroupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, verbose: true }) });

    assert(loggerLogSpy.calledOnceWith(calendarGroupResponse));
    assert(logToStderrSpy.calledOnce);
  });

  it('retrieves calendar group by name for the signed-in user with verbose output', async () => {
    const logToStderrSpy = sinon.spy(logger, 'logToStderr');
    const expectedFilterUrl = `https://graph.microsoft.com/v1.0/me/calendarGroups?$filter=name eq 'Personal%20Events'`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === expectedFilterUrl) {
        return calendarGroupsResponseForFilter;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, verbose: true }) });

    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponseForFilter.value[0]));
    assert(logToStderrSpy.calledOnceWith(`Retrieving calendar group '${calendarGroupName}'...`));
  });

  it('throws an error when calendar group name does not match any results', async () => {
    const expectedFilterUrl = `https://graph.microsoft.com/v1.0/me/calendarGroups?$filter=name eq 'Personal%20Events'`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === expectedFilterUrl) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName }) }),
      new CommandError(`The specified calendar group '${calendarGroupName}' does not exist.`)
    );
  });

  it('retrieves calendar group for a user specified by userName using delegated permissions with shared scope (ReadWrite.Shared)', async () => {
    sinonUtil.restore(accessToken.getScopesFromAccessToken);
    sinon.stub(accessToken, 'getScopesFromAccessToken').returns(['Calendars.ReadWrite.Shared']);

    const encodedUserName = formatting.encodeQueryParameter(userName);
    const expectedFilterUrl = `https://graph.microsoft.com/v1.0/users('${encodedUserName}')/calendarGroups?$filter=name eq 'Personal%20Events'`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === expectedFilterUrl) {
        return calendarGroupsResponseForFilter;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userName }) });
    assert(loggerLogSpy.calledOnceWith(calendarGroupsResponseForFilter.value[0]));
  });
});

