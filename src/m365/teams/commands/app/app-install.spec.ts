import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { settingsNames } from '../../../../settingsNames.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-install.js';

describe(commands.APP_INSTALL, () => {
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
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_INSTALL);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when neither teamId, userId nor userName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when teamId and userId are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        teamId: '00000000-0000-0000-0000-000000000000',
        userId: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when teamId and userName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        teamId: '00000000-0000-0000-0000-000000000000',
        userName: 'steve@contoso.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userId and userName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        userId: '00000000-0000-0000-0000-000000000000',
        userName: 'steve@contoso.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
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
        teamId: '00000000-0000-0000-0000-000000000000'
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
        teamId: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails to get Teams app when app does not exist', async () => {
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
        teamId: '00000000-0000-0000-0000-000000000000'
      }
    } as any), new CommandError('The specified Teams app does not exist'));
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        id: 'not-a78e-fd77-4599-97a5-dbb6372846c5',
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        userId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
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
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
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

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: '5b31c38c-2584-42f0-aa47-657fb3a84230' });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/c527a470-a882-481c-981c-ee6efaba85c7/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/5b31c38c-2584-42f0-aa47-657fb3a84230"}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        name: 'Test app'
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('passes validation when the id and teamId are correct', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id and userId are correct', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id and userName are correct', async () => {
    const actual = await command.validate({
      options: {
        id: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        userName: 'steve@contoso.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds app from the catalog to a Microsoft Team', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/c527a470-a882-481c-981c-ee6efaba85c7/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('adds app from the catalog to a Microsoft Team by name (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/appCatalogs/teamsApps?$filter=displayName eq '`) > -1) {
        return {
          "value": [
            {
              "id": "4440558e-8c73-4597-abc7-3644a64c4bce",
              "displayName": "Test app"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/c527a470-a882-481c-981c-ee6efaba85c7/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        name: 'Test app',
        debug: true
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('installs app from the catalog the user specified with userId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7`) {
        return {
          "value": [
            {
              "businessPhones": [
                "425-555-0100"
              ],
              "displayName": "MOD Administrator",
              "givenName": "MOD",
              "jobTitle": null,
              "mail": "admin@contoso.OnMicrosoft.com",
              "mobilePhone": "425-555-0101",
              "officeLocation": null,
              "preferredLanguage": "en-US",
              "surname": "Administrator",
              "userPrincipalName": "admin@contoso.onmicrosoft.com",
              "id": "c527a470-a882-481c-981c-ee6efaba85c7"
            }
          ]
        };
      }

      throw 'Invalid request';
    });
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
    });
    assert.strictEqual(log.length, 0);
  });

  it('installs app from the catalog the user specified with userId (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7`) {
        return {
          "value": [
            {
              "businessPhones": [
                "425-555-0100"
              ],
              "displayName": "MOD Administrator",
              "givenName": "MOD",
              "jobTitle": null,
              "mail": "admin@contoso.OnMicrosoft.com",
              "mobilePhone": "425-555-0101",
              "officeLocation": null,
              "preferredLanguage": "en-US",
              "surname": "Administrator",
              "userPrincipalName": "admin@contoso.onmicrosoft.com",
              "id": "c527a470-a882-481c-981c-ee6efaba85c7"
            }
          ]
        };
      }

      throw 'Invalid request';
    });
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
        id: '4440558e-8c73-4597-abc7-3644a64c4bce',
        debug: true
      }
    });
  });

  it('installs app from the catalog the user specified with userName', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/steve%40contoso.com/teamwork/installedApps` &&
        JSON.stringify(opts.data) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        userName: 'steve@contoso.com',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('correctly handles error while installing Teams app', async () => {
    const error = {
      "error": {
        "code": "UnKnown",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T12:14:15",
          "request-id": "1d6fc213-9f35-4cb3-b496-3d8b10aebdfa",
          "client-request-id": "1d6fc213-9f35-4cb3-b496-3d8b10aebdfa"
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    } as any), new CommandError(error.error.message));
  });

  it(`correctly handles error when trying to install an app for a user that doesn't exist (invalid user name)`, async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "NotFound",
          "message": "Failed to find user with id 'steve@contoso.com' in the tenant",
          "innerError": {
            "date": "2022-02-14T12:14:15",
            "request-id": "1d6fc213-9f35-4cb3-b496-3d8b10aebdfa",
            "client-request-id": "1d6fc213-9f35-4cb3-b496-3d8b10aebdfa"
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { userName: 'steve@contoso.com', id: '4440558e-8c73-4597-abc7-3644a64c4bce' } } as any), new CommandError("Failed to find user with id 'steve@contoso.com' in the tenant"));
  });

  it(`correctly handles error when trying to install an app for a user that doesn't exist (invalid user ID)`, async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7`) {
        throw {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource 'c527a470-a882-481c-981c-ee6efaba85c7' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2022-02-14T13:27:37",
              "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
              "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
            }
          }
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').rejects('Invalid request');

    await assert.rejects(command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    } as any), new CommandError("User with ID c527a470-a882-481c-981c-ee6efaba85c7 not found. Original error: Resource 'c527a470-a882-481c-981c-ee6efaba85c7' does not exist or one of its queried reference-property objects are not present."));
  });

  it(`correctly handles error when trying to install an app for a user that doesn't exist (invalid user ID; debug)`, async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7`) {
        throw {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource 'c527a470-a882-481c-981c-ee6efaba85c7' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2022-02-14T13:27:37",
              "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
              "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
            }
          }
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').rejects('Invalid request');

    await assert.rejects(command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: '4440558e-8c73-4597-abc7-3644a64c4bce',
        debug: true
      }
    } as any), new CommandError("User with ID c527a470-a882-481c-981c-ee6efaba85c7 not found. Original error: Resource 'c527a470-a882-481c-981c-ee6efaba85c7' does not exist or one of its queried reference-property objects are not present."));
  });
});
