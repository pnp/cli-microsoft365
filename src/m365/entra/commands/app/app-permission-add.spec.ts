import { Application, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-permission-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_PERMISSION_ADD, () => {
  const appId = '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d';
  const appName = 'My App';
  const appObjectId = '2aaf2d9e-815e-4a3e-bb80-9b3d9c79078b';
  const servicePrincipalId = '7c330108-8825-4b6c-b280-8d1d68da6bd7';
  const servicePrincipals: ServicePrincipal[] = [{ "appId": appId, 'id': servicePrincipalId, "servicePrincipalNames": [] }, { "appId": "00000003-0000-0000-c000-000000000000", "id": "fb4be1df-eaa6-4bd0-a068-71f9b2cbe2be", "servicePrincipalNames": ["https://canary.graph.microsoft.com/", "https://graph.microsoft.us/", "https://dod-graph.microsoft.us/", "00000003-0000-0000-c000-000000000000/ags.windows.net", "00000003-0000-0000-c000-000000000000", "https://canary.graph.microsoft.com", "https://graph.microsoft.com", "https://ags.windows.net", "https://graph.microsoft.us", "https://graph.microsoft.com/", "https://dod-graph.microsoft.us"], "appRoles": [{ "allowedMemberTypes": ["Application"], "description": "Allows the app to read and update user profiles without a signed in user.", "displayName": "Read and write all users' full profiles", "id": "741f803b-c850-494e-b5df-cde7c675a1ca", "isEnabled": true, "origin": "Application", "value": "User.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read user profiles without a signed in user.", "displayName": "Read all users' full profiles", "id": "df021288-bdef-4463-88db-98f22de89214", "isEnabled": true, "origin": "Application", "value": "User.Read.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read and query your audit log activities, without a signed-in user.", "displayName": "Read all audit log data", "id": "b0afded3-3588-46d8-8b3d-9842eff778da", "isEnabled": true, "origin": "Application", "value": "AuditLog.Read.All" }], "oauth2PermissionScopes": [{ "adminConsentDescription": "Allows the app to see and update the data you gave it access to, even when users are not currently using the app. This does not give the app any additional permissions.", "adminConsentDisplayName": "Maintain access to data you have given it access to", "id": "7427e0e9-2fba-42fe-b0c0-848c9e6a8182", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to see and update the data you gave it access to, even when you are not currently using the app. This does not give the app any additional permissions.", "userConsentDisplayName": "Maintain access to data you have given it access to", "value": "offline_access" }, { "adminConsentDescription": "Allows the app to read the available Teams templates, on behalf of the signed-in user.", "adminConsentDisplayName": "Read available Teams templates", "id": "cd87405c-5792-4f15-92f7-debc0db6d1d6", "isEnabled": true, "type": "User", "userConsentDescription": "Read available Teams templates, on your behalf.", "userConsentDisplayName": "Read available Teams templates", "value": "TeamTemplates.Read" }] }];
  const applications: Application[] = [{ 'id': appObjectId, 'appId': appId, 'requiredResourceAccess': [] }];
  const multipleApplications: Application[] = [{ 'id': appObjectId, 'appId': appId, 'requiredResourceAccess': [] }, { 'id': '2aaf2d9e-815e-4a3e-bb80-9b3d9c79078c', 'appId': '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3e', 'requiredResourceAccess': [] }];
  const applicationPermissions = 'https://graph.microsoft.com/User.ReadWrite.All https://graph.microsoft.com/User.Read.All';
  const delegatedPermissions = 'https://graph.microsoft.com/offline_access';

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      request.post,
      odata.getAllItems,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_PERMISSION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds application permissions to app specified by appObjectId without granting admin consent', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=id eq '${appObjectId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${applications[0].id}`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appObjectId: appObjectId, applicationPermissions: applicationPermissions, verbose: true } });
    assert(patchStub.called);
  });

  it('adds application permissions to app specified by appName without granting admin consent', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq 'My%20App'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${applications[0].id}`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: appName, applicationPermissions: applicationPermissions, verbose: true } });
    assert(patchStub.called);
  });

  it('adds application permissions to app specified by appId without granting admin consent', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${applications[0].id}`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appId: appId, applicationPermissions: applicationPermissions, verbose: true } });
    assert(patchStub.called);
  });

  it('adds application permissions to app specified by appId while granting admin consent', async () => {
    let amountOfPostCalls = 0;
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${applications[0].id}`) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments`) {
        amountOfPostCalls += 1;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appId: appId, applicationPermissions: applicationPermissions, grantAdminConsent: true, verbose: true } });
    assert.strictEqual(amountOfPostCalls, 2);
  });

  it('creates service principal if not exists before granting admin consent', async () => {
    let numberOfPostCalls = 0;
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return [servicePrincipals[1]];
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${applications[0].id}`) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/servicePrincipals') {
        numberOfPostCalls++;
        return { "appId": appId, 'id': servicePrincipalId, "servicePrincipalNames": [] };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments`) {
        numberOfPostCalls++;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appId: appId, applicationPermissions: applicationPermissions, grantAdminConsent: true, verbose: true } });
    assert.strictEqual(numberOfPostCalls, 3);
  });

  it('adds delegated and application permissions to appId while granting admin consent', async () => {
    let amountOfPostCalls = 0;

    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${applications[0].id}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/oauth2PermissionGrants`) {
        amountOfPostCalls++;
        return;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments`) {
        amountOfPostCalls++;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appId: appId, delegatedPermissions: delegatedPermissions, applicationPermissions: applicationPermissions, grantAdminConsent: true, verbose: true } });
    assert.strictEqual(amountOfPostCalls, 3);
  });

  it('throws an error when application specified by appId is not found', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return [];
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { appId: appId, applicationPermissions: applicationPermissions, verbose: true } }),
      new CommandError(`App with client id ${appId} not found in Microsoft Entra ID`));
  });

  it('throws an error when application specified by appObjectId is not found', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=id eq '${appObjectId}'&$select=id,appId,requiredResourceAccess`:
          return [];
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { appObjectId: appObjectId, applicationPermissions: applicationPermissions, verbose: true } }),
      new CommandError(`App with object id ${appObjectId} not found in Microsoft Entra ID`));
  });

  it('throws an error when service principal is not found', async () => {
    const api = 'https://grax.microsoft.com/User.ReadWrite.All';
    const pos: number = api.lastIndexOf('/');
    const servicePrincipalName: string = api.substring(0, pos);
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { appId: appId, applicationPermissions: api, verbose: true } }),
      new CommandError(`Service principal ${servicePrincipalName} not found`));
  });

  it('throws an error when permission is not found', async () => {
    const api = 'https://graph.microsoft.com/NotFound.All';
    const pos: number = api.lastIndexOf('/');
    const servicePrincipalName: string = api.substring(0, pos);
    const permissionName: string = api.substring(pos + 1);
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,appId,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { appId: appId, applicationPermissions: api, verbose: true } }),
      new CommandError(`Permission ${permissionName} for service principal ${servicePrincipalName} not found`));
  });

  it('handles error when multiple apps with the specified name found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq 'My%20App'&$select=id,appId,requiredResourceAccess`:
          return [{ 'id': '9b1b1e42-794b-4c71-93ac-5ed92488b67f', 'appId': appId, 'requiredResourceAccess': [] }, { 'id': '9b1b1e42-794b-4c71-93ac-5ed92488b67g', 'appId': appId, 'requiredResourceAccess': [] }];
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: appName
      }
    }), new CommandError(`Multiple Entra application registrations with name 'My App' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`));
  });

  it('handles a non-existent app by appName', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq 'My%20App'&$select=id,appId,requiredResourceAccess`:
          return [];
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { appName: appName } }),
      new CommandError(`App with name ${appName} not found in Microsoft Entra ID`));
  });

  it('handles selecting single result when multiple apps with the specified name found and cli is set to prompt', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq 'My%20App'&$select=id,appId,requiredResourceAccess`:
          return multipleApplications;
        default:
          throw 'Invalid request';
      }
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(multipleApplications[0]);

    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${multipleApplications[0].id}`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: appName, applicationPermissions: applicationPermissions, verbose: true } });
    assert(patchStub.called);
  });

  it('passes validation if applicationPermission is passed', async () => {
    const actual = await command.validate({ options: { appId: appId, applicationPermissions: applicationPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if delegatedPermission is passed', async () => {
    const actual = await command.validate({ options: { appId: appId, delegatedPermissions: delegatedPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if both applicationPermission or delegatedPermission are passed', async () => {
    const actual = await command.validate({ options: { appId: appId, applicationPermissions: applicationPermissions, delegatedPermissions: delegatedPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if both applicationPermission or delegatedPermission is not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: '123', applicationPermissions: applicationPermissions } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the appId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: appId, applicationPermissions: applicationPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the appObjectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: '123', applicationPermissions: applicationPermissions } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the appObjectId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: appObjectId, applicationPermissions: applicationPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if neither the appId nor the appObjectId are provided.', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        applicationPermissions: applicationPermissions
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both appId and appObjectId are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        appId: appId,
        appObjectId: appObjectId,
        applicationPermissions: applicationPermissions
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});