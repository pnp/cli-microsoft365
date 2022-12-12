import * as assert from 'assert';
import { AxiosRequestConfig } from 'axios';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./task-set');

describe(commands.TASK_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let patchStub: sinon.SinonStub<[options: AxiosRequestConfig<any>]>;

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
  const patchRequestData = {
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
    patchStub = sinon.stub(request, 'patch').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==/tasks/abc`) {
        return Promise.resolve(patchRequestData);
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
      request.patch,
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
    assert.strictEqual(command.name.startsWith(commands.TASK_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates tasks for  list using listId', async () => {
    await command.action(logger, {
      options: {
        id: 'abc',
        title: "New task",
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=='
      }
    } as any);
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(patchRequestData));
  });

  it('updates tasks for list using listName (debug)', async () => {
    await command.action(logger, {
      options: {
        id: 'abc',
        title: "New task",
        listName: 'Tasks List',
        status: "notStarted",
        debug: true
      }
    } as any);
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(patchRequestData));
  });

  it('updates tasks for list with bodyContent and bodyContentType', async () => {
    const bodyText = '<h3>Lorem ipsum</h3>';
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        bodyContent: bodyText,
        bodyContentType: 'html'
      }
    } as any);

    assert.strictEqual(patchStub.lastCall.args[0].data.body.content, bodyText);
    assert.strictEqual(patchStub.lastCall.args[0].data.body.contentType, 'html');
  });

  it('updates tasks for list with bodyContent and no bodyContentType', async () => {
    const bodyText = 'Lorem ipsum';
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        bodyContent: bodyText
      }
    } as any);

    assert.strictEqual(patchStub.lastCall.args[0].data.body.content, bodyText);
    assert.strictEqual(patchStub.lastCall.args[0].data.body.contentType, 'text');
  });

  it('updates tasks for list with importance', async () => {
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        importance: 'high'
      }
    } as any);

    assert.strictEqual(patchStub.lastCall.args[0].data.importance, 'high');
  });

  it('updates tasks for list with dueDateTime', async () => {
    const dateTime = '2023-01-01';
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        dueDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(patchStub.lastCall.args[0].data.dueDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });

  it('updates tasks for list with reminderDateTime', async () => {
    const dateTime = '2023-01-01T12:00:00';
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        reminderDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(patchStub.lastCall.args[0].data.reminderDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
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
    
    await assert.rejects(command.action(logger, { options: {
      id: 'abc',
      title: "New task",
      listName: 'Tasks List',
      debug: true } } as any), new CommandError('The specified task list does not exist'));
  });

  it('handles error correctly', async () => {
    sinonUtil.restore(request.patch);
    sinon.stub(request, 'patch').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: {
      listName: "Tasks List",
      id: 'abc',
      title: "New task" } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if both listId and listName options are passed', async () => {
    const actual = await command.validate({
      options: {
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        listName: 'Tasks List',
        title: 'New Task',
        id: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither listId nor listName options are passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New Task',
        id: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id not passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New Task',
        listName: 'Tasks List'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if status is not allowed value', async () => {
    const options: any = {
      title: 'New Task',
      id: 'abc',
      status: "test",
      listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=='

    };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, 'test is not a valid value. Allowed values are notStarted|inProgress|completed|waitingOnOthers|deferred');
  });

  it('fails validation when invalid bodyContentType is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        id: 'abc',
        status: "notStarted",
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
        id: 'abc',
        status: "notStarted",
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
        id: 'abc',
        status: "notStarted",
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
        id: 'abc',
        status: "notStarted",
        listName: 'Tasks List',
        reminderDateTime: '01/01/2022'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly validates the arguments', async () => {
    const actual = await command.validate({
      options: {
        title: 'New Task',
        id: 'abc',
        status: "notStarted",
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=='
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });


  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
