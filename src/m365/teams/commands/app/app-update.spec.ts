import assert from 'assert';
import fs from 'fs';
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
import command from './app-update.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_UPDATE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.put,
      fs.readFileSync,
      fs.existsSync,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_UPDATE);
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
        name: 'Test app',
        filePath: 'teamsapp.zip'
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
        filePath: 'teamsapp.zip'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID.', async () => {
    const actual = await command.validate({
      options: {
        id: 'invalid',
        filePath: 'teamsapp.zip'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the filePath does not exist', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = await command.validate({
      options: { id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22", filePath: 'invalid.zip' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the filePath points to a directory', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').returns(true);
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(stats);

    const actual = await command.validate({
      options: { id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22", filePath: './' }
    }, commandInfo);
    sinonUtil.restore([
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').returns(false);
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(stats);

    const actual = await command.validate({
      options: {
        id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
        filePath: 'teamsapp.zip'
      }
    }, commandInfo);
    sinonUtil.restore([
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('fails to get Teams app when app does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Test app',
        filePath: 'teamsapp.zip'
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
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
              "displayName": "Test app"
            },
            {
              "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
              "displayName": "Test app"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Test app',
        filePath: 'teamsapp.zip'
      }
    } as any), new CommandError('Multiple Teams apps with name Test app found. Found: e3e29acb-8c79-412b-b746-e6c39ff4cd22, 5b31c38c-2584-42f0-aa47-657fb3a84230.'));
  });

  it('handles selecting single result when multiple Teams apps found with the specified name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
              "displayName": "Test app"
            },
            {
              "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
              "displayName": "Test app"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves({ id: '5b31c38c-2584-42f0-aa47-657fb3a84230' });

    let updateTeamsAppCalled = false;
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/5b31c38c-2584-42f0-aa47-657fb3a84230`) {
        updateTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    await command.action(logger, { options: { filePath: 'teamsapp.zip', name: 'Test app' } });
    assert(updateTeamsAppCalled);
  });

  it('update Teams app in the tenant app catalog by id', async () => {
    let updateTeamsAppCalled = false;
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        updateTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    await command.action(logger, { options: { filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } });
    assert(updateTeamsAppCalled);
  });

  it('update Teams app in the tenant app catalog by id (debug)', async () => {
    let updateTeamsAppCalled = false;

    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        updateTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    await command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } });
    assert(updateTeamsAppCalled);
  });

  it('update Teams app in the tenant app catalog by name (debug)', async () => {
    let updateTeamsAppCalled = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
              "displayName": "Test app"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        updateTeamsAppCalled = true;
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    await command.action(logger, {
      options: {
        debug: true,
        filePath: 'teamsapp.zip',
        name: 'Test app'
      }
    });
    assert(updateTeamsAppCalled);
  });

  it('correctly handles error when updating an app', async () => {
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
    sinon.stub(request, 'put').rejects(error);

    sinon.stub(fs, 'readFileSync').returns('123');

    await assert.rejects(command.action(logger, { options: { filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } } as any), new CommandError('An error has occurred'));
  });
});
