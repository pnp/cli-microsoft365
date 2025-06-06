import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './permission-list.js';
import { appRegApplicationPermissions, appRegDelegatedPermissionsMultipleResources, appRegNoApiPermissions, flowServiceOAuth2PermissionScopes, msGraphPrincipalAppRoles, msGraphPrincipalOAuth2PermissionScopes } from './permission-list.mock.js';

describe(commands.PERMISSION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  //#region Mocked Responses 
  const appId = '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d';
  const appResponse = {
    "apps": [
      {
        "appId": appId,
        "name": "CLI app1"
      }
    ]
  };
  //#endregion

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify(appResponse));
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      entraApp.getAppRegistrationByAppId
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PERMISSION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves permissions from app registration if service principal not found', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appRegDelegatedPermissionsMultipleResources);

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/appRoles`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/appRoles`:
          return { value: [] };
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return appRegDelegatedPermissionsMultipleResources;
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '7df0a125-d3be-4c96-aa54-591f83ff541c'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "7df0a125-d3be-4c96-aa54-591f83ff541c",
                "id": "582d24e0-4dd7-41c5-b7dd-2a52817a95aa",
                "displayName": "Microsoft Flow Service"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
                "id": "c7c82441-65de-4fb1-ac2e-83a947ced55f",
                "displayName": "Windows Azure Service Management API"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/oauth2PermissionScopes`:
          return flowServiceOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/oauth2PermissionScopes`:
          return {
            "value": [
              {
                "adminConsentDescription": "Allows the application to access the Azure Management Service API acting as users in the organization.",
                "adminConsentDisplayName": "Access Azure Service Management as organization users (preview)",
                "id": "41094075-9dad-400e-a0bd-54e686782033",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": "Allows the application to access Azure Service Management as you.",
                "userConsentDisplayName": "Access Azure Service Management as you (preview)",
                "value": "user_impersonation"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "Microsoft Flow Service",
        "permission": "Flows.Read.All",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Flow Service",
        "permission": "Flows.Manage.All",
        "type": "Delegated"
      },
      {
        "resource": "Windows Azure Service Management API",
        "permission": "user_impersonation",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "AccessReview.Read.All",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "Agreement.Read.All",
        "type": "Delegated"
      }
    ]));
  });

  it('retrieves permissions from app registration if service principal not found (debug)', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appRegDelegatedPermissionsMultipleResources);

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/appRoles`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/appRoles`:
          return { value: [] };
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return appRegDelegatedPermissionsMultipleResources;
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '7df0a125-d3be-4c96-aa54-591f83ff541c'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "7df0a125-d3be-4c96-aa54-591f83ff541c",
                "id": "582d24e0-4dd7-41c5-b7dd-2a52817a95aa",
                "displayName": "Microsoft Flow Service"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
                "id": "c7c82441-65de-4fb1-ac2e-83a947ced55f",
                "displayName": "Windows Azure Service Management API"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/oauth2PermissionScopes`:
          return flowServiceOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/oauth2PermissionScopes`:
          return {
            "value": [
              {
                "adminConsentDescription": "Allows the application to access the Azure Management Service API acting as users in the organization.",
                "adminConsentDisplayName": "Access Azure Service Management as organization users (preview)",
                "id": "41094075-9dad-400e-a0bd-54e686782033",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": "Allows the application to access Azure Service Management as you.",
                "userConsentDisplayName": "Access Azure Service Management as you (preview)",
                "value": "user_impersonation"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) });
    assert(loggerLogToStderrSpy.called);
  });

  it('retrieves delegated permissions from app registration', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appRegDelegatedPermissionsMultipleResources);

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/appRoles`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/appRoles`:
          return { value: [] };
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return appRegDelegatedPermissionsMultipleResources;
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '7df0a125-d3be-4c96-aa54-591f83ff541c'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "7df0a125-d3be-4c96-aa54-591f83ff541c",
                "id": "582d24e0-4dd7-41c5-b7dd-2a52817a95aa",
                "displayName": "Microsoft Flow Service"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
                "id": "c7c82441-65de-4fb1-ac2e-83a947ced55f",
                "displayName": "Windows Azure Service Management API"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/oauth2PermissionScopes`:
          return flowServiceOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/oauth2PermissionScopes`:
          return {
            "value": [
              {
                "adminConsentDescription": "Allows the application to access the Azure Management Service API acting as users in the organization.",
                "adminConsentDisplayName": "Access Azure Service Management as organization users (preview)",
                "id": "41094075-9dad-400e-a0bd-54e686782033",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": "Allows the application to access Azure Service Management as you.",
                "userConsentDisplayName": "Access Azure Service Management as you (preview)",
                "value": "user_impersonation"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "Microsoft Flow Service",
        "permission": "Flows.Read.All",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Flow Service",
        "permission": "Flows.Manage.All",
        "type": "Delegated"
      },
      {
        "resource": "Windows Azure Service Management API",
        "permission": "user_impersonation",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "AccessReview.Read.All",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "Agreement.Read.All",
        "type": "Delegated"
      }
    ]));
  });

  it('retrieves application permissions from app registration', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appRegApplicationPermissions);

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "Microsoft Graph",
        "permission": "AppCatalog.Read.All",
        "type": "Application"
      }
    ]));
  });

  it(`doesn't fail when the app registration has no API permissions`, async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appRegNoApiPermissions);

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return Promise.resolve({ value: [] });
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([]));
  });

  it('retrieves permissions for a service principal with delegated and app permissions', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return {
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return {
            "value": [
              {
                "clientId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "consentType": "AllPrincipals",
                "id": "UWHjFHLkzk6BLD6AqD-j9Y2z0olAbOtGs5b239cP8Hw",
                "principalId": null,
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "scope": "Mail.Read offline_access"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return {
            "value": [
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9UpXcjOo6AhAhgmM8i3vZlM",
                "deletedDateTime": null,
                "appRoleId": "01d4889c-1287-42c6-ac1f-5d1e02578ef6",
                "createdDateTime": "2021-11-21T19:03:41.5442462Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              },
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9WuT_ApPC4hHr5iOlpdxK_M",
                "deletedDateTime": null,
                "appRoleId": "810c84a8-4a9e-49e6-bf7d-12d183f40d01",
                "createdDateTime": "2021-11-21T19:03:41.63799Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return {
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "Microsoft Graph",
        "permission": "Files.Read.All",
        "type": "Application"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "Mail.Read",
        "type": "Application"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "Mail.Read",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "offline_access",
        "type": "Delegated"
      }
    ]));
  });

  it('retrieves permissions for a service principal with delegated and app permissions (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return {
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return {
            "value": [
              {
                "clientId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "consentType": "AllPrincipals",
                "id": "UWHjFHLkzk6BLD6AqD-j9Y2z0olAbOtGs5b239cP8Hw",
                "principalId": null,
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "scope": "Mail.Read offline_access"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return {
            "value": [
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9UpXcjOo6AhAhgmM8i3vZlM",
                "deletedDateTime": null,
                "appRoleId": "01d4889c-1287-42c6-ac1f-5d1e02578ef6",
                "createdDateTime": "2021-11-21T19:03:41.5442462Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              },
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9WuT_ApPC4hHr5iOlpdxK_M",
                "deletedDateTime": null,
                "appRoleId": "810c84a8-4a9e-49e6-bf7d-12d183f40d01",
                "createdDateTime": "2021-11-21T19:03:41.63799Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return {
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) });
    assert(loggerLogToStderrSpy.called);
  });

  it('retrieves permissions for a service principal with delegated permissions', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return {
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return {
            "value": [
              {
                "clientId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "consentType": "AllPrincipals",
                "id": "UWHjFHLkzk6BLD6AqD-j9Y2z0olAbOtGs5b239cP8Hw",
                "principalId": null,
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "scope": "Mail.Read offline_access"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return { "value": [] };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return {
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "Microsoft Graph",
        "permission": "Mail.Read",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "offline_access",
        "type": "Delegated"
      }
    ]));
  });

  it('handles non-existent app', async () => {
    const error = `App with appId '${appId}' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppId').rejects(new Error(error));

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return { value: [] };
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError(error));
  });

  it('handles error when retrieving service principal for the Microsoft Entra app', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          throw {
            error: {
              message: `An error has occurred`
            }
          };
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError(`An error has occurred`));
  });

  it('handles error when retrieving OAuth2 permission grants for service principal', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return {
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          throw {
            error: {
              message: `An error has occurred`
            }
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return { value: [] };
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError(`An error has occurred`));
  });

  it('handles error when retrieving OAuth2 permission scopes for service principal', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return {
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return {
            "value": [
              {
                "clientId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "consentType": "AllPrincipals",
                "id": "UWHjFHLkzk6BLD6AqD-j9Y2z0olAbOtGs5b239cP8Hw",
                "principalId": null,
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "scope": "Mail.Read offline_access"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return {
            "value": [
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9UpXcjOo6AhAhgmM8i3vZlM",
                "deletedDateTime": null,
                "appRoleId": "01d4889c-1287-42c6-ac1f-5d1e02578ef6",
                "createdDateTime": "2021-11-21T19:03:41.5442462Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              },
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9WuT_ApPC4hHr5iOlpdxK_M",
                "deletedDateTime": null,
                "appRoleId": "810c84a8-4a9e-49e6-bf7d-12d183f40d01",
                "createdDateTime": "2021-11-21T19:03:41.63799Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return {
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          throw {
            error: {
              message: `An error has occurred`
            }
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError(`An error has occurred`));
  });

  it('handles error when retrieving app role assignments for service principal', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return {
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return {
            "value": [
              {
                "clientId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "consentType": "AllPrincipals",
                "id": "UWHjFHLkzk6BLD6AqD-j9Y2z0olAbOtGs5b239cP8Hw",
                "principalId": null,
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "scope": "Mail.Read offline_access"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          throw {
            error: {
              message: `An error has occurred`
            }
          };
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError(`An error has occurred`));
  });

  it('handles error when retrieving app roles for service principal', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return {
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return {
            "value": [
              {
                "clientId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "consentType": "AllPrincipals",
                "id": "UWHjFHLkzk6BLD6AqD-j9Y2z0olAbOtGs5b239cP8Hw",
                "principalId": null,
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "scope": "Mail.Read offline_access"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return {
            "value": [
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9UpXcjOo6AhAhgmM8i3vZlM",
                "deletedDateTime": null,
                "appRoleId": "01d4889c-1287-42c6-ac1f-5d1e02578ef6",
                "createdDateTime": "2021-11-21T19:03:41.5442462Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              },
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9WuT_ApPC4hHr5iOlpdxK_M",
                "deletedDateTime": null,
                "appRoleId": "810c84a8-4a9e-49e6-bf7d-12d183f40d01",
                "createdDateTime": "2021-11-21T19:03:41.63799Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return {
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          throw {
            error: {
              message: `An error has occurred`
            }
          };
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError(`An error has occurred`));
  });

  it('handles error when retrieving Microsoft Entra registration', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          throw {
            error: {
              message: `An error has occurred`
            }
          };
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError(`An error has occurred`));
  });

  it('handles non-existent service principal from app registration permissions', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appRegApplicationPermissions);

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return { value: [] };
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "00000003-0000-0000-c000-000000000000",
        "permission": "e12dae10-5a57-4817-b79d-dfbec5348930",
        "type": "Application"
      }
    ]));
  });

  it('handles service principal referencing a non-existent app permission', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return {
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return {
            "value": [
              {
                "clientId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "consentType": "AllPrincipals",
                "id": "UWHjFHLkzk6BLD6AqD-j9Y2z0olAbOtGs5b239cP8Hw",
                "principalId": null,
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "scope": "Mail.Read offline_access"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return {
            "value": [
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9UpXcjOo6AhAhgmM8i3vZlM",
                "deletedDateTime": null,
                "appRoleId": "01d4889c-1287-42c6-ac1f-5d1e02578ef7",
                "createdDateTime": "2021-11-21T19:03:41.5442462Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              },
              {
                "id": "UWHjFHLkzk6BLD6AqD-j9WuT_ApPC4hHr5iOlpdxK_M",
                "deletedDateTime": null,
                "appRoleId": "810c84a8-4a9e-49e6-bf7d-12d183f40d01",
                "createdDateTime": "2021-11-21T19:03:41.63799Z",
                "principalDisplayName": "CLI app",
                "principalId": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "principalType": "ServicePrincipal",
                "resourceDisplayName": "Microsoft Graph",
                "resourceId": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return {
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "Microsoft Graph",
        "permission": "01d4889c-1287-42c6-ac1f-5d1e02578ef7",
        "type": "Application"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "Mail.Read",
        "type": "Application"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "Mail.Read",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "offline_access",
        "type": "Delegated"
      }
    ]));
  });

  it('handles unknown delegated permissions from app registration', async () => {
    const appReg = appRegDelegatedPermissionsMultipleResources;
    appReg.requiredResourceAccess[0].resourceAccess[0].id = "e45c5562-459d-4d1b-8148-83eb1b6dcf84";
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appReg);

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/appRoles`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/appRoles`:
          return { value: [] };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '7df0a125-d3be-4c96-aa54-591f83ff541c'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "7df0a125-d3be-4c96-aa54-591f83ff541c",
                "id": "582d24e0-4dd7-41c5-b7dd-2a52817a95aa",
                "displayName": "Microsoft Flow Service"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
                "id": "c7c82441-65de-4fb1-ac2e-83a947ced55f",
                "displayName": "Windows Azure Service Management API"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/oauth2PermissionScopes`:
          return flowServiceOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/oauth2PermissionScopes`:
          return {
            "value": [
              {
                "adminConsentDescription": "Allows the application to access the Azure Management Service API acting as users in the organization.",
                "adminConsentDisplayName": "Access Azure Service Management as organization users (preview)",
                "id": "41094075-9dad-400e-a0bd-54e686782033",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": "Allows the application to access Azure Service Management as you.",
                "userConsentDisplayName": "Access Azure Service Management as you (preview)",
                "value": "user_impersonation"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "Microsoft Flow Service",
        "permission": "e45c5562-459d-4d1b-8148-83eb1b6dcf84",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Flow Service",
        "permission": "Flows.Manage.All",
        "type": "Delegated"
      },
      {
        "resource": "Windows Azure Service Management API",
        "permission": "user_impersonation",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "AccessReview.Read.All",
        "type": "Delegated"
      },
      {
        "resource": "Microsoft Graph",
        "permission": "Agreement.Read.All",
        "type": "Delegated"
      }
    ]));
  });

  it('handles unknown application permissions from app registration', async () => {
    const appReg = appRegApplicationPermissions;
    appReg.requiredResourceAccess[0].resourceAccess[0].id = 'e12dae10-5a57-4817-b79d-dfbec5348931';
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appReg);

    sinon.stub(request, 'get').callsFake(async opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=appId,id,displayName`:
          return { value: [] };
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return {
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          };
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return msGraphPrincipalOAuth2PermissionScopes;
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return msGraphPrincipalAppRoles;
        default:
          throw `Invalid request ${JSON.stringify(opts)}`;
      }
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
      {
        "resource": "Microsoft Graph",
        "permission": "e12dae10-5a57-4817-b79d-dfbec5348931",
        "type": "Application"
      }
    ]));
  });
});