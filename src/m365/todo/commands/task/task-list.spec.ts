import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './task-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TASK_LIST, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'status', 'createdDateTime', 'lastModifiedDateTime']);
  });

  it('fails validation if both listId and listName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        listId: 'AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA==',
        listName: 'Tasks List'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither listId nor listName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails to get ToDo Task list when the specified task list does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/me/todo/lists?$filter=displayName eq '`) > -1) {
        return { value: [] };
      }
      throw 'The specified task list does not exist';
    });

    await assert.rejects(command.action(logger, { options: { listName: 'Tasks List' } } as any), new CommandError('The specified task list does not exist'));
  });

  it('passes validation if only listId is passed', async () => {
    const actual = await command.validate({
      options: {
        listId: 'AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA=='
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only listName is passed', async () => {
    const actual = await command.validate({
      options: {
        listName: 'Tasks List'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists To Do tasks using listId in JSON output mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/tasks`) > -1) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        listId: "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA=="
      }
    });
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
  });

  it('lists To Do tasks using listName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/me/todo/lists?$filter=displayName eq '`) > -1) {
        return {
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
        };
      }

      if ((opts.url as string).indexOf(`/tasks`) > -1) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        listName: 'Tasks List',
        output: 'text'
      }
    });
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
  });
});
