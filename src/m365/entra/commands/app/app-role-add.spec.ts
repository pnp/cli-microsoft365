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
import command from './app-role-add.js';
import { settingsNames } from '../../../../settingsNames.js';
import { entraApp } from '../../../../utils/entraApp.js';

describe(commands.APP_ROLE_ADD, () => {
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      entraApp.getAppRegistrationByAppId,
      entraApp.getAppRegistrationByAppName,
      entraApp.getAppRegistrationByObjectId
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_ROLE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates app role for the specified appId, app has no roles', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves({
      id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      appRoles: []
    });

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.displayName === 'Role' &&
          appRole.description === 'Custom role' &&
          appRole.value === 'Custom.Role' &&
          JSON.stringify(appRole.allowedMemberTypes) === JSON.stringify(['User'])) {
          return;
        }
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    });
  });

  it('creates app role for the specified appObjectId, app has one role', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByObjectId').resolves({
      id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      appRoles: [{
        "allowedMemberTypes": [
          "User"
        ],
        "description": "Managers",
        "displayName": "Managers",
        "id": "c4352a0a-494f-46f9-b843-479855c173a7",
        "isEnabled": true,
        "origin": "Application",
        "value": "managers"
      }]
    });

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[1];
        if (JSON.stringify({
          "allowedMemberTypes": [
            "User"
          ],
          "description": "Managers",
          "displayName": "Managers",
          "id": "c4352a0a-494f-46f9-b843-479855c173a7",
          "isEnabled": true,
          "origin": "Application",
          "value": "managers"
        }) === JSON.stringify(opts.data.appRoles[0]) &&
          appRole.displayName === 'Role' &&
          appRole.description === 'Custom role' &&
          appRole.value === 'Custom.Role' &&
          JSON.stringify(appRole.allowedMemberTypes) === JSON.stringify(['Application'])) {
          return;
        }
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'applications',
        claim: 'Custom.Role',
        verbose: true
      }
    });
  });

  it('creates app role for the specified appName, app has multiple roles', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppName').resolves({
      id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      appRoles: [
        {
          "allowedMemberTypes": [
            "User"
          ],
          "description": "Managers",
          "displayName": "Managers",
          "id": "c4352a0a-494f-46f9-b843-479855c173a7",
          "isEnabled": true,
          "origin": "Application",
          "value": "managers"
        },
        {
          "allowedMemberTypes": [
            "User"
          ],
          "description": "Team leads",
          "displayName": "Team leads",
          "id": "c4352a0a-494f-46f9-b843-479855c173a8",
          "isEnabled": true,
          "origin": "Application",
          "value": "teamLeads"
        }
      ]
    });
    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 3) {
        const appRole = opts.data.appRoles[2];
        if (appRole.displayName === 'Role' &&
          appRole.description === 'Custom role' &&
          appRole.value === 'Custom.Role' &&
          JSON.stringify(appRole.allowedMemberTypes) === JSON.stringify(['User', 'Application'])) {
          return;
        }
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appName: 'My app',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'both',
        claim: 'Custom.Role',
        verbose: true
      }
    });
  });

  it('handles error when the app specified with appObjectId not found', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByObjectId').throws({
      "error": {
        "code": "Request_ResourceNotFound",
        "message": "Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.",
        "innerError": {
          "date": "2021-04-20T17:22:30",
          "request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d",
          "client-request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d"
        }
      }
    });
    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }), new CommandError(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('handles error when the app specified with the appId not found', async () => {
    const error = `App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppId').rejects(new Error(error));

    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }), new CommandError(`App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`));
  });

  it('handles error when the app specified with appName not found', async () => {
    const error = `App with name 'My app' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));
    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }), new CommandError(error));
  });

  it('handles error when multiple apps with the specified appName found', async () => {
    const error = `Multiple apps with name 'My app' found in Microsoft Entra ID. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));

    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    }), new CommandError(error));
  });

  it('handles error when retrieving information about app through appName failed', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));
    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('handles error when retrieving app roles failed', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByObjectId').throws({
      "error": {
        "code": "Request_ResourceNotFound",
        "message": "Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.",
        "innerError": {
          "date": "2021-04-20T17:22:30",
          "request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d",
          "client-request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d"
        }
      }
    });
    sinon.stub(request, 'patch').rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    } as any), new CommandError(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('handles error when updating app roles failed', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByObjectId').resolves({
      id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      appRoles: [{
        "allowedMemberTypes": [
          "User"
        ],
        "description": "Managers",
        "displayName": "Managers",
        "id": "c4352a0a-494f-46f9-b843-479855c173a7",
        "isEnabled": true,
        "origin": "Application",
        "value": "managers"
      }]
    });
    sinon.stub(request, 'patch').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'Role',
        description: 'Custom role',
        allowedMembers: 'usersGroups',
        claim: 'Custom.Role'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if appId and appObjectId specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appObjectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', name: 'Managers', description: 'Managers', allowedMembers: 'userGroups', claim: 'managers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app', name: 'Managers', description: 'Managers', allowedMembers: 'userGroups', claim: 'managers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app', name: 'Managers', description: 'Managers', allowedMembers: 'userGroups', claim: 'managers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, appObjectId nor appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { name: 'Managers', description: 'Managers', allowedMembers: 'userGroups', claim: 'managers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid allowedMembers specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowedMembers: 'invalid', name: 'Managers', description: 'Managers', claim: 'managers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if claim length exceeds 120 chars', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowedMembers: 'usersGroups', claim: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Cras ullamcorper, arcu vel finibus facilisis, orci velit lectus.', name: 'Managers', description: 'Managers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if claim starts with a .', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowedMembers: 'usersGroups', claim: '.claim', name: 'Managers', description: 'Managers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if claim contains invalid characters', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', allowedMembers: 'usersGroups', claim: 'cláim', name: 'Managers', description: 'Managers' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (appId)', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'Role', description: 'Custom role', allowedMembers: 'usersGroups', claim: 'Custom.Role' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (appObjectId)', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'Role', description: 'Custom role', allowedMembers: 'usersGroups', claim: 'Custom.Role' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (appName)', async () => {
    const actual = await command.validate({ options: { appName: 'My app', name: 'Role', description: 'Custom role', allowedMembers: 'usersGroups', claim: 'Custom.Role' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('returns an empty array for an invalid member type', () => {
    const actual = (command as any).getAllowedMemberTypes({ options: { allowedMembers: 'foo' } });
    assert.deepStrictEqual(actual, []);
  });
});
