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
import command, { options } from './calendar-remove.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';
import { calendar } from '../../../../utils/calendar.js';

describe(commands.CALENDAR_REMOVE, () => {
  const userId = 'ae0e8388-cd70-427f-9503-c57498ee3337';
  const userName = 'john.doe@contoso.com';
  const calendarId = 'AAMkADJmMVAAA=';
  const calendarName = 'Volunteer';
  const calendarGroupId = 'AQMkADJmMVAAA=';
  const calendarGroupName = 'My Calendars';

  let log: any[];
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

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.post,
      calendar.getUserCalendarByName,
      calendarGroup.getUserCalendarGroupByName,
      cli.handleMultipleResultsFound,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALENDAR_REMOVE);
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

  it('removes the calendar by id for a user specified by id without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: calendarId,
        userId: userId,
        force: true,
        verbose: true
      }) });
    assert(deleteRequestStub.called);
  });

  it('permanently removes the calendar by id for a user specified by id without prompting for confirmation', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}/permanentDelete`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: calendarId,
        userId: userId,
        permanent: true,
        force: true,
        verbose: true
      })
    });
    assert(postRequestStub.called);
  });

  it('removes the calendar by id for a user specified by name from a calendar group specified by name while prompting for confirmation', async () => {
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').withArgs(userName, calendarGroupName, 'id').resolves({ id: calendarGroupId });
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups/${calendarGroupId}/calendars/${calendarId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: calendarId,
        userName: userName,
        calendarGroupName: calendarGroupName,
        verbose: true
      }) });
    assert(deleteRequestStub.called);
  });

  it('removes the calendar by name for a user specified by name from a calendar group specified by name while prompting for confirmation', async () => {
    sinon.stub(calendar, 'getUserCalendarByName').withArgs(userName, calendarName, calendarGroupId, 'id').resolves({ id: calendarId });
    sinon.stub(calendarGroup, 'getUserCalendarGroupByName').withArgs(userName, calendarGroupName, 'id').resolves({ id: calendarGroupId });
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups/${calendarGroupId}/calendars/${calendarId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        name: calendarName,
        userName: userName,
        calendarGroupName: calendarGroupName,
        verbose: true
      })
    });
    assert(deleteRequestStub.called);
  });

  it('removes the calendar by name for a user specified by name from a calendar group specified by id without prompting for confirmation', async () => {
    sinon.stub(calendar, 'getUserCalendarByName').withArgs(userName, calendarName, calendarGroupId, 'id').resolves({ id: calendarId });
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/calendarGroups/${calendarGroupId}/calendars/${calendarId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        name: calendarName,
        userName: userName,
        calendarGroupId: calendarGroupId,
        force: true,
        verbose: true
      }) });
    assert(deleteRequestStub.called);
  });

  it('prompts before removing the calendar when confirm option not passed', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: calendarId,
        userId: userId
      })
    });

    assert(promptIssued);
  });

  it('aborts removing the calendar when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: calendarId,
        userId: userId
      })
    });
    assert(deleteSpy.notCalled);
  });

  it('throws an error when the calendar specified by id for a user specified by id cannot be found', async () => {
    const error = {
      error: {
        code: 'ErrorItemNotFound',
        message: 'The specified object was not found in the store.'
      }
    };
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/calendars/${calendarId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        id: calendarId,
        userId: userId,
        force: true
      })
    }), new CommandError(error.error.message));
  });
});
