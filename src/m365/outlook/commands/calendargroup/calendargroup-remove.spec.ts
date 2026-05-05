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
import command, { options } from './calendargroup-remove.js';

describe(commands.CALENDARGROUP_REMOVE, () => {
  const calendarGroupId = 'AAMkAGE0MGM1Y2M5LWEzMmUtNGVlNy05MjRlLTk0YmYyY2I5NTM3ZAAuAAAAAAC_0WfqSjt_SqLtNkuO-bj1AQAbfYq5lmBxQ6a4t1fGbeYAAAAAAEOAAA=';
  const calendarGroupName = 'Personal Events';
  const currentUserId = '2b4097f3-5b17-4153-a8b4-cd680e333555';
  const userId = 'b743445a-112c-4fda-9afd-05943f9c7b36';
  const userName = 'john.doe@contoso.com';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean;

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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(accessToken, 'getUserIdFromAccessToken').returns(currentUserId);
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      accessToken.getUserIdFromAccessToken,
      calendarGroup.getUserCalendarGroupByName,
      request.get,
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDARGROUP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when neither id nor name is specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and name are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, name: calendarGroupName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both userId and userName are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, userId, userName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, userId: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, userName: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with id', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with name', () => {
    const actual = commandOptionsSchema.safeParse({ name: calendarGroupName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id and userId', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with id and userName', () => {
    const actual = commandOptionsSchema.safeParse({ id: calendarGroupId, userName });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing when force option not passed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId }) });

    assert(promptIssued);
  });

  it('aborts removing when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId }) });
    assert(deleteSpy.notCalled);
  });

  it('removes the calendar group specified by id for the signed-in user without prompting', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${currentUserId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, force: true }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes the calendar group specified by id for the signed-in user (verbose)', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${currentUserId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, force: true, verbose: true }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes the calendar group specified by name for the signed-in user', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').resolves({ id: calendarGroupId, name: calendarGroupName });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${currentUserId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes the calendar group specified by name for the signed-in user (verbose)', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').resolves({ id: calendarGroupId, name: calendarGroupName });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${currentUserId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, force: true, verbose: true }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes the calendar group specified by id for a user specified by userId', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, userId, force: true }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes the calendar group specified by id for a user specified by userName', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('john.doe%40contoso.com')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, userName, force: true }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes the calendar group specified by name for a user specified by userId', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').resolves({ id: calendarGroupId, name: calendarGroupName });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userId, force: true }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes the calendar group specified by name for a user specified by userName', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').resolves({ id: calendarGroupId, name: calendarGroupName });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('john.doe%40contoso.com')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, userName, force: true }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('removes the calendar group specified by id using app-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendarGroups/${calendarGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, userId, force: true }) });
    assert(deleteRequestStub.calledOnce);
  });

  it('throws error when running with app-only permissions without userId or userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, force: true }) }),
      new CommandError('When running with application permissions either userId or userName is required.')
    );
  });

  it('throws error when calendar group specified by name is not found', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').rejects(new Error(`The specified calendar group '${calendarGroupName}' does not exist.`));

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ name: calendarGroupName, force: true }) }),
      new CommandError(`The specified calendar group '${calendarGroupName}' does not exist.`)
    );
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = `Your request can't be completed. The calendar group '${calendarGroupName}' is not empty.`;
    sinon.stub(request, 'delete').rejects({ error: { error: { code: 'ErrorInvalidRequest', message: errorMessage } } });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({ id: calendarGroupId, force: true }) }),
      new CommandError(errorMessage)
    );
  });
});
