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
import command from './app-role-list.js';
import { settingsNames } from '../../../../settingsNames.js';
import { entraApp } from '../../../../utils/entraApp.js';

describe(commands.APP_ROLE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  //#region Mocked Responses 
  const appResponse = {
    value: [
      {
        "id": "5b31c38c-2584-42f0-aa47-657fb3a84230"
      }
    ]
  };
  //#endregion

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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      entraApp.getAppRegistrationByAppId,
      entraApp.getAppRegistrationByAppName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_ROLE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['displayName', 'description', 'id']);
  });

  it('lists roles for the specified appId (debug)', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return {
          value: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Readers",
              "displayName": "Readers",
              "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
              "isEnabled": true,
              "origin": "Application",
              "value": "readers"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Writers",
              "displayName": "Writers",
              "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
              "isEnabled": true,
              "origin": "Application",
              "value": "writers"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8' } });
    assert(loggerLogSpy.calledWith([
      {
        "allowedMemberTypes": [
          "User"
        ],
        "description": "Readers",
        "displayName": "Readers",
        "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
        "isEnabled": true,
        "origin": "Application",
        "value": "readers"
      },
      {
        "allowedMemberTypes": [
          "User"
        ],
        "description": "Writers",
        "displayName": "Writers",
        "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
        "isEnabled": true,
        "origin": "Application",
        "value": "writers"
      }
    ]));
  });

  it('lists roles for the specified appName (debug)', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppName').resolves(appResponse.value[0]);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return {
          value: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Readers",
              "displayName": "Readers",
              "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
              "isEnabled": true,
              "origin": "Application",
              "value": "readers"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Writers",
              "displayName": "Writers",
              "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
              "isEnabled": true,
              "origin": "Application",
              "value": "writers"
            }
          ]
        };
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { debug: true, appName: 'My app' } });
    assert(loggerLogSpy.calledWith([
      {
        "allowedMemberTypes": [
          "User"
        ],
        "description": "Readers",
        "displayName": "Readers",
        "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
        "isEnabled": true,
        "origin": "Application",
        "value": "readers"
      },
      {
        "allowedMemberTypes": [
          "User"
        ],
        "description": "Writers",
        "displayName": "Writers",
        "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
        "isEnabled": true,
        "origin": "Application",
        "value": "writers"
      }
    ]));
  });

  it('lists roles for the specified appObjectId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return {
          value: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Readers",
              "displayName": "Readers",
              "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
              "isEnabled": true,
              "origin": "Application",
              "value": "readers"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Writers",
              "displayName": "Writers",
              "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
              "isEnabled": true,
              "origin": "Application",
              "value": "writers"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230' } });
    assert(loggerLogSpy.calledWith([
      {
        "allowedMemberTypes": [
          "User"
        ],
        "description": "Readers",
        "displayName": "Readers",
        "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
        "isEnabled": true,
        "origin": "Application",
        "value": "readers"
      },
      {
        "allowedMemberTypes": [
          "User"
        ],
        "description": "Writers",
        "displayName": "Writers",
        "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
        "isEnabled": true,
        "origin": "Application",
        "value": "writers"
      }
    ]));
  });

  it(`returns an empty array if the specified app has no roles`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230' } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('handles error when the app specified with appObjectId not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles') {
        throw {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2021-04-20T17:22:30",
              "request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d",
              "client-request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d"
            }
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230'
      }
    }), new CommandError(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('handles error when the app specified with the appId not found', async () => {
    const error = `App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppId').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }), new CommandError(`App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`));
  });

  it('handles error when the app specified with appName not found', async () => {
    const error = `App with name 'My app' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app'
      }
    }), new CommandError(error));
  });

  it('handles error when multiple apps with the specified appName found', async () => {
    const error = `Multiple apps with name 'My app' found in Microsoft Entra ID. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app'
      }
    }), new CommandError(error));
  });

  it('handles error when retrieving information about app through appName failed', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app'
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

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appObjectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, appObjectId nor appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if appId specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if appObjectId specified', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if appName specified', async () => {
    const actual = await command.validate({ options: { appName: 'My app' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
