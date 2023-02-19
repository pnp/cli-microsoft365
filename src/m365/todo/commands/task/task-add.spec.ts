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
    "importance": "low",
    "isReminderOn": false,
    "status": "notStarted",
    "title": "New task",
    "createdDateTime": "2020-10-28T10:30:20.9783659Z",
    "lastModifiedDateTime": "2020-10-28T10:30:21.3616801Z",
    "id": "AAMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwBGAAAAAAAq-A2AAw08T7MU1EldWtTXBwCEcpBkci0MQ5UKfxTfG4DYAAGZB5U-AACEcpBkci0MQ5UKfxTfG4DYAAGhnfKPAAA=",
    "body": {
      "content": "",
      "contentType": "text"
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    postStub = sinon.stub(request, 'post').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==/tasks`) {
        return Promise.resolve(postRequestData);
      }
      return Promise.reject();
    });

    sinon.stub(request, 'get').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks%20List'`) {
        return Promise.resolve(getRequestData);
      }
      return Promise.reject();
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
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
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

  it('rejects if no tasks list is found with the specified list name', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks%20List'`) {
        return Promise.resolve(
          {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('4cb2b035-ad76-406c-bdc4-6c72ad403a22')/todo/lists",
            "value": []
          }
        );
      }
      return Promise.reject();
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
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

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
