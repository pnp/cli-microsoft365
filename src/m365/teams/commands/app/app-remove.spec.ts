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
import command from './app-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let requests: any[];
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
    requests = [];
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and name options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'e3e29acb-8c79-412b-b746-e6c39ff4cd22',
        name: 'TeamsApp'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name options are not passed', async () => {
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

  it('fails validation if the id is not a valid GUID.', async () => {
    const actual = await command.validate({
      options: { id: 'invalid' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('removes Teams app by id in the tenant app catalog with confirmation (debug)', async () => {
    let removeTeamsAppCalled = false;
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, force: true } });
    assert(removeTeamsAppCalled);
  });

  it('removes Teams app by id in the tenant app catalog without confirmation', async () => {
    let removeTeamsAppCalled = false;
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } });
    assert(removeTeamsAppCalled);
  });

  it('aborts removing Teams app when prompt not confirmed', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    command.action(logger, { options: { id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } });
    assert(requests.length === 0);
  });

  it('removes Teams app by name in the tenant app catalog without confirmation (debug)', async () => {
    let removeTeamsAppCalled = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'&$select=id`) {
        return {
          "value": [
            {
              "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
              "displayName": "TeamsApp"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.doesNotReject(command.action(logger, { options: { debug: true, name: 'TeamsApp' } }));
    assert(removeTeamsAppCalled);
  });

  it('handles selecting single result when multiple teams apps to remove with the specified name are found and cli is set to prompt', async () => {
    let removeTeamsAppCalled = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'&$select=id`) {
        return {
          "value": [
            { "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22" },
            { "id": "9b1b1e42-794b-4c71-93ac-5ed92488b67g" }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: 'e3e29acb-8c79-412b-b746-e6c39ff4cd22' });
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.doesNotReject(command.action(logger, { options: { debug: true, name: 'TeamsApp' } }));
    assert(removeTeamsAppCalled);
  });

  it('fails to get Teams app when app does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'&$select=id`) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'TeamsApp',
        force: true
      }
    } as any), new CommandError('The specified Teams app does not exist'));
  });

  it('handles error when multiple Teams apps with the specified name found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'&$select=id`) {
        return {
          "value": [
            {
              "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
              "displayName": "TeamsApp"
            },
            {
              "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
              "displayName": "TeamsApp"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'TeamsApp',
        force: true
      }
    } as any), new CommandError(`Multiple Teams apps with name 'TeamsApp' found. Found: e3e29acb-8c79-412b-b746-e6c39ff4cd22, 5b31c38c-2584-42f0-aa47-657fb3a84230.`));
  });

  it('correctly handles error when removing app', async () => {
    sinon.stub(request, 'delete').rejects({
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    });
    await assert.rejects(command.action(logger, {
      options: {
        filePath: 'teamsapp.zip',
        id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, force: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
