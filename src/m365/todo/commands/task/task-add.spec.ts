import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request, { CliRequestOptions } from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./task-add');

describe(commands.TASK_ADD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let postStub: sinon.SinonStub<[options: CliRequestOptions]>;

  const getRequestData = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('4cb2b035-ad76-406c-bdc4-6c72ad403a22')/todo/lists",
    "value": [
      {
        "@odata.etag": "W/\"hHKQZHItDEOVCn8U3xuA2AABoWDAng==\"",
        "displayName": "Tasks List",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "none",
        "id": "AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=="
      }
    ]
  };
  const postRequestData = {
    title: "New task",
    "body": {
      contentType: "text"
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    (command as any).items = [];
    postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==/tasks`) {
        return postRequestData;
      }
      throw null;
    });

    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks%20List'`) {
        return getRequestData;
      }
      throw null;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Date.now
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds To Do task to task list using listId', async () => {
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=='
      }
    } as any);
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(postRequestData));
  });

  it('adds To Do task to task list using listName (debug)', async () => {
    await command.action(logger, {
      options: {
        title: 'New task',
        listName: 'Tasks List',
        debug: true
      }
    } as any);
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(postRequestData));
  });

  it('adds To Do task with bodyContent and bodyContentType', async () => {
    const bodyText = 'Lorem ipsum';
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        bodyContent: bodyText,
        bodyContentType: 'text'
      }
    } as any);

    assert.strictEqual(postStub.lastCall.args[0].data.body.content, bodyText);
    assert.strictEqual(postStub.lastCall.args[0].data.body.contentType, 'text');
  });

  it('adds To Do task with importance', async () => {
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        importance: 'high'
      }
    } as any);

    assert.strictEqual(postStub.lastCall.args[0].data.importance, 'high');
  });

  it('adds To Do task with dueDateTime', async () => {
    const dateTime = '2023-01-01';
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        dueDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(postStub.lastCall.args[0].data.dueDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });

  it('adds To Do task with reminderDateTime', async () => {
    const dateTime = '2023-01-01T12:00:00';
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        reminderDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(postStub.lastCall.args[0].data.reminderDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });

  it('adds To Do task with categories ', async () => {
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        categories: 'None,Preset24'
      }
    } as any);

    assert.deepStrictEqual(postStub.lastCall.args[0].data.categories, ['None', 'Preset24']);
  });

  it('adds To Do task with completedDateTime', async () => {
    const dateTime = '2023-01-01';
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        completedDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(postStub.lastCall.args[0].data.completedDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });

  it('adds To Do task with startDateTime', async () => {
    const dateTime = '2023-01-01';
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        startDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(postStub.lastCall.args[0].data.startDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });

  it('adds To Do task with status', async () => {
    await command.action(logger, {
      options: {
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: 'inProgress'
      }
    } as any);

    assert.deepStrictEqual(postStub.lastCall.args[0].data.status, 'inProgress');
  });

  it('rejects if no tasks list is found with the specified list name', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks%20List'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('4cb2b035-ad76-406c-bdc4-6c72ad403a22')/todo/lists",
          "value": []
        };
      }
      throw null;
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: 'New task',
        listName: 'Tasks List',
        debug: true
      }
    } as any), new CommandError('The specified task list does not exist'));
  });

  it('handles error correctly', async () => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { listName: "Tasks List", title: "New task" } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation when invalid bodyContentType is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        bodyContentType: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid importance is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        importance: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid dueDateTime is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        dueDateTime: '01/01/2022'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid reminderDateTime is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        reminderDateTime: '01/01/2022'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid completedDateTime is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        completedDateTime: '01/01/2022',
        status: 'completed'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when valid completedDateTime is passed without status completed', async () => {
    const dateTime = '2023-01-01';
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        completedDateTime: dateTime,
        status: 'inProgress'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when valid completedDateTime is passed without status', async () => {
    const dateTime = '2023-01-01';
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        completedDateTime: dateTime
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid startDateTime is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        startDateTime: '01/01/2022'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid status is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        listName: 'Tasks List',
        status: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if listId and title options are passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New Task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=='
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
