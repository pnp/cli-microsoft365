import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./task-list');

describe(commands.TASK_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'status', 'createdDateTime', 'lastModifiedDateTime']);
  });

  it('fails validation if both listId and listName options are passed', () => {
    const actual = command.validate({
      options: {
        listId: 'AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA==',
        listName: 'Tasks List'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither listId nor listName options are passed', () => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails to get ToDo Task list when the specified task list does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/me/todo/lists?$filter=displayName eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('The specified task list does not exist');
    });

    command.action(logger, {
      options: {
        debug: true,
        listName: 'Tasks List'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified task list does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation if only listId is passed', () => {
    const actual = command.validate({
      options: {
        listId: 'AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA=='
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if only listName is passed', () => {
    const actual = command.validate({
      options: {
        listName: 'Tasks List'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('lists To Do tasks using listId (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/tasks`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('e1251b10-1ba4-49e3-b35a-933e3f21772b')/todo/lists('AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA%3D%3D')/tasks",
          "value": [
            {
              "@odata.etag": "W/\"xMBBaLl1lk+dAn8KkjfXKQABF7wl/A==\"",
              "importance": "normal",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Stay healthy",
              "createdDateTime": "2020-11-01T17:13:13.9582172Z",
              "lastModifiedDateTime": "2020-11-01T17:13:15.1645231Z",
              "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=",
              "body": {
                "content": "",
                "contentType": "text"
              }
            },
            {
              "@odata.etag": "W/\"xMBBaLl1lk+dAn8KkjfXKQABF7wl8w==\"",
              "importance": "normal",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Eat food",
              "createdDateTime": "2020-11-01T17:13:10.7970391Z",
              "lastModifiedDateTime": "2020-11-01T17:13:13.1037095Z",
              "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuOAAA=",
              "body": {
                "content": "",
                "contentType": "text"
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        listId: "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA=="
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWithExactly(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists To Do tasks using listId in JSON output mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/tasks`) > -1) {
        return Promise.resolve({
          value: [
            {
              "importance": "normal",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Stay healthy",
              "createdDateTime": "2020-11-01T17:13:13.9582172Z",
              "lastModifiedDateTime": "2020-11-01T17:13:15.1645231Z",
              "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=",
              "body": {
                "content": "",
                "contentType": "text"
              }
            },
            {
              "importance": "normal",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Eat food",
              "createdDateTime": "2020-11-01T17:13:10.7970391Z",
              "lastModifiedDateTime": "2020-11-01T17:13:13.1037095Z",
              "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuOAAA=",
              "body": {
                "content": "",
                "contentType": "text"
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        listId: "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA=="
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [
            {
              "importance": "normal",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Stay healthy",
              "createdDateTime": "2020-11-01T17:13:13.9582172Z",
              "lastModifiedDateTime": "2020-11-01T17:13:15.1645231Z",
              "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=",
              "body": {
                "content": "",
                "contentType": "text"
              }
            },
            {
              "importance": "normal",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Eat food",
              "createdDateTime": "2020-11-01T17:13:10.7970391Z",
              "lastModifiedDateTime": "2020-11-01T17:13:13.1037095Z",
              "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuOAAA=",
              "body": {
                "content": "",
                "contentType": "text"
              }
            }
          ]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists To Do tasks using listName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/me/todo/lists?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('e1251b10-1ba4-49e3-b35a-933e3f21772b')/todo/lists",
          "value": [
            {
              "@odata.etag": "W/\"xMBBaLl1lk+dAn8KkjfXKQABF7wr8w==\"",
              "displayName": "Tasks List",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA=="
            }
          ]
        });
      }

      if ((opts.url as string).indexOf(`/tasks`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('e1251b10-1ba4-49e3-b35a-933e3f21772b')/todo/lists('AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA%3D%3D')/tasks",
          "value": [
            {
              "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=",
              "title": "Stay healthy",
              "status": "notStarted",
              "createdDateTime": "2020-11-01T17:13:13.9582172Z",
              "lastModifiedDateTime": "2020-11-01T17:13:15.1645231Z"
            },
            {
              "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuOAAA=",
              "title": "Eat food",
              "status": "notStarted",
              "createdDateTime": "2020-11-01T17:13:10.7970391Z",
              "lastModifiedDateTime": "2020-11-01T17:13:13.1037095Z"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        listName: 'Tasks List'
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify([
          {
            "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=",
            "title": "Stay healthy",
            "status": "notStarted",
            "createdDateTime": "2020-11-01T17:13:13.9582172Z",
            "lastModifiedDateTime": "2020-11-01T17:13:15.1645231Z"
          },
          {
            "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuOAAA=",
            "title": "Eat food",
            "status": "notStarted",
            "createdDateTime": "2020-11-01T17:13:10.7970391Z",
            "lastModifiedDateTime": "2020-11-01T17:13:13.1037095Z"
          }
        ]);
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});