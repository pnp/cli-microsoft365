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
import command from './user-app-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.USER_APP_ADD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_APP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the userId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        userId: 'invalid',
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userName is not a valid UPN.', async () => {
    const actual = await command.validate({
      options: {
        userName: "not-a-upn",
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        id: 'not-c49b-4fd4-8223-28f0ac3a6402',
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id and name are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: '15d7a78e-fd77-4599-97a5-dbb6372846c5', name: 'TeamsApp', userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither id nor name are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the input is correct', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the input is correct (userName)', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        userName: 'admin@contoso.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds app by id from the catalog for the specified user by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    } as any);
  });

  it('adds app by id from the catalog for the specified user by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/admin%40contoso.com/teamwork/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        userName: 'admin@contoso.com',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    } as any);
  });

  it('adds app by name from the catalog for the specified user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'`) {
        return {
          "value": [
            {
              "id": "YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=",
              "displayName": "TeamsApp"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY="}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        name: 'TeamsApp'
      }
    } as any);
  });

  it('fails to get teams app when app by name does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        name: 'TeamsApp'
      }
    } as any), new CommandError('The specified Teams app does not exist'));
  });

  it('handles selecting single result when multiple teams apps with the specified name found and cli is set to prompt', async () => {
    let addRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'`) {
        return {
          "value": [
            {
              "id": "ZDczZWVjZmQtYzFkNS00MzY2LWJkMjEtZDUyOTM1ZThkYjkxIyMxLjYuMC4wIyNQdWJsaXNoZWQ=",
              "displayName": "TeamsApp"
            },
            {
              "id": "NmY0ODM2N2EtMjVmMC00NjNmLTlmMGQtMmFiZTBiYmYzNzRjIyMxLjAuMCMjUHVibGlzaGVk",
              "displayName": "TeamsApp"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: "ZDczZWVjZmQtYzFkNS00MzY2LWJkMjEtZDUyOTM1ZThkYjkxIyMxLjYuMC4wIyNQdWJsaXNoZWQ=" });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/ZDczZWVjZmQtYzFkNS00MzY2LWJkMjEtZDUyOTM1ZThkYjkxIyMxLjYuMC4wIyNQdWJsaXNoZWQ="}`) {
        addRequestIssued = true;
        return Promise.resolve();
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userId: 'c527a470-a882-481c-981c-ee6efaba85c7', name: 'TeamsApp' } });
    assert(addRequestIssued);
  });

  it('correctly handles error while installing teams app', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };

    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
