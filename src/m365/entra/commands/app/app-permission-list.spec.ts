import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-permission-list.js';
import { settingsNames } from '../../../../settingsNames.js';
import { application, graphApplication, graphOauth2PermissionScope, spOnlineApplication, spOnlineOauth2PermissionScope, allPermissionsResponse, applicationPermissionsResponse, delegatedPermissionsResponse, applicationWithoutPermissions, applicationWithUnknownPermissions, allUnknownPermissionsResponse, allUnkownServicePrincipalPermissionsResponse } from './app-permission-list.mock.js';
import { CommandError } from '../../../../Command.js';
import { entraApp } from '../../../../utils/entraApp.js';

describe(commands.APP_PERMISSION_LIST, () => {
  const appId = '2bf26ae1-9be3-425f-a393-5fe8390e3a36';
  const appName = 'My App';
  const appObjectId = '29807f3b-fef6-4985-b987-8c2565d021bc';

  //#region Mocked Responses 
  const appResponse = {
    value: [
      {
        "id": appObjectId
      }
    ]
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
    assert.strictEqual(command.name, commands.APP_PERMISSION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the appId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: appId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the appObjectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the appObjectId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: appObjectId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if neither the appId, appName, nor appObjectId are provided.', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        type: 'all'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when appId, appName, and appObjectId are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        appId: appId,
        appName: appName,
        appObjectId: appObjectId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the type is not a valid permission type', async () => {
    const actual = await command.validate({ options: { appId: appId, type: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('lists the permissions of an app registration when using objectId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appObjectId}`) {
        return application;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0ff1-ce00-000000000000",
              "id": "5d72c3ba-e836-4be3-94fb-fa6057b1611b",
              "displayName": "Office 365 SharePoint Online"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0000-c000-000000000000",
              "id": "6aac2819-1b16-4d85-be7b-4bc1d1a456a7",
              "displayName": "Microsoft Graph"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/6aac2819-1b16-4d85-be7b-4bc1d1a456a7/oauth2PermissionScopes`) {
        return spOnlineOauth2PermissionScope;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/6aac2819-1b16-4d85-be7b-4bc1d1a456a7/appRoles`) {
        return spOnlineApplication;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/5d72c3ba-e836-4be3-94fb-fa6057b1611b/oauth2PermissionScopes`) {
        return graphOauth2PermissionScope;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/5d72c3ba-e836-4be3-94fb-fa6057b1611b/appRoles`) {
        return graphApplication;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, { options: { appObjectId: appObjectId, verbose: true } });
    assert(loggerLogSpy.calledWith(allPermissionsResponse));
  });

  it('lists the application permissions of an app registration when using objectId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appObjectId}`) {
        return application;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0ff1-ce00-000000000000",
              "id": "5d72c3ba-e836-4be3-94fb-fa6057b1611b",
              "displayName": "Office 365 SharePoint Online"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0000-c000-000000000000",
              "id": "6aac2819-1b16-4d85-be7b-4bc1d1a456a7",
              "displayName": "Microsoft Graph"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/6aac2819-1b16-4d85-be7b-4bc1d1a456a7/appRoles`) {
        return spOnlineApplication;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/5d72c3ba-e836-4be3-94fb-fa6057b1611b/appRoles`) {
        return graphApplication;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { appObjectId: appObjectId, type: 'application' } });
    assert(loggerLogSpy.calledWith(applicationPermissionsResponse));
  });

  it('lists the delegated permissions of an app registration when using appId', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '2bf26ae1-9be3-425f-a393-5fe8390e3a36'&$select=id`) {
        return { value: [{ id: '29807f3b-fef6-4985-b987-8c2565d021bc' }] };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appObjectId}`) {
        return application;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0ff1-ce00-000000000000",
              "id": "5d72c3ba-e836-4be3-94fb-fa6057b1611b",
              "displayName": "Office 365 SharePoint Online"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0000-c000-000000000000",
              "id": "6aac2819-1b16-4d85-be7b-4bc1d1a456a7",
              "displayName": "Microsoft Graph"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/6aac2819-1b16-4d85-be7b-4bc1d1a456a7/oauth2PermissionScopes`) {
        return spOnlineOauth2PermissionScope;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/5d72c3ba-e836-4be3-94fb-fa6057b1611b/oauth2PermissionScopes`) {
        return graphOauth2PermissionScope;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { appId: appId, type: 'delegated', debug: true } });
    assert(loggerLogSpy.calledWith(delegatedPermissionsResponse));
  });

  it('lists the delegated permissions of an app registration when using appName', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppName').resolves(appResponse.value[0]);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appObjectId}`) {
        return application;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0ff1-ce00-000000000000",
              "id": "5d72c3ba-e836-4be3-94fb-fa6057b1611b",
              "displayName": "Office 365 SharePoint Online"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0000-c000-000000000000",
              "id": "6aac2819-1b16-4d85-be7b-4bc1d1a456a7",
              "displayName": "Microsoft Graph"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/6aac2819-1b16-4d85-be7b-4bc1d1a456a7/oauth2PermissionScopes`) {
        return spOnlineOauth2PermissionScope;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/5d72c3ba-e836-4be3-94fb-fa6057b1611b/oauth2PermissionScopes`) {
        return graphOauth2PermissionScope;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { appName: appName, type: 'delegated', verbose: true } });
    assert(loggerLogSpy.calledWith(delegatedPermissionsResponse));
  });

  it('handles error when multiple apps with the specified name found', async () => {
    const error = `Multiple apps with name 'My app' found in Microsoft Entra ID. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        appName: appName
      }
    }), new CommandError(error));
  });

  it('handles a non-existent app by appId', async () => {
    const error = `App with appId '${appId}' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppId').rejects(new Error(error));

    await assert.rejects(command.action(logger, { options: { appId: appId } }),
      new CommandError(error));
  });

  it('handles a non-existent app by appName', async () => {
    const error = `App with name 'My app' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));

    await assert.rejects(command.action(logger, { options: { appName: appName } }),
      new CommandError(error));
  });

  it('lists no permissions for app registration without permissions', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appObjectId}`) {
        return applicationWithoutPermissions;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { appObjectId: appObjectId } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('handles unknown permissions from app registration', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appObjectId}`) {
        return applicationWithUnknownPermissions;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0ff1-ce00-000000000000",
              "id": "5d72c3ba-e836-4be3-94fb-fa6057b1611b",
              "displayName": "Office 365 SharePoint Online"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": [
            {
              "appId": "00000003-0000-0000-c000-000000000000",
              "id": "6aac2819-1b16-4d85-be7b-4bc1d1a456a7",
              "displayName": "Microsoft Graph"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/6aac2819-1b16-4d85-be7b-4bc1d1a456a7/oauth2PermissionScopes`) {
        return spOnlineOauth2PermissionScope;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/6aac2819-1b16-4d85-be7b-4bc1d1a456a7/appRoles`) {
        return spOnlineApplication;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/5d72c3ba-e836-4be3-94fb-fa6057b1611b/oauth2PermissionScopes`) {
        return graphOauth2PermissionScope;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/5d72c3ba-e836-4be3-94fb-fa6057b1611b/appRoles`) {
        return graphApplication;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { appObjectId: appObjectId } });
    assert(loggerLogSpy.calledWith(allUnknownPermissionsResponse));
  });

  it('handles unknown service principal from app registration using object id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appObjectId}`) {
        return application;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": []
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`) {
        return {
          "value": []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { appObjectId: appObjectId } });
    assert(loggerLogSpy.calledWith(allUnkownServicePrincipalPermissionsResponse));
  });

  it('handles error when retrieving Entra app registration', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appObjectId}`) {
        return application;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`) {
        throw {
          error: {
            message: `An error has occurred`
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, { options: { appObjectId: appObjectId } }), new CommandError(`An error has occurred`));
  });

  it('handles error when retrieving Entra app registration using name', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app'
      }
    } as any), new CommandError('An error has occurred'));
  });
});