import * as assert from 'assert';
import * as sinon from 'sinon';
import * as fs from 'fs';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
import { appRegApplicationPermissions, appRegDelegatedPermissionsMultipleResources, appRegNoApiPermissions, flowServiceOAuth2PermissionScopes, msGraphPrincipalAppRoles, msGraphPrincipalOAuth2PermissionScopes } from './permission-list.mock';
const command: Command = require('./permission-list');

describe(commands.PERMISSION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      "apps": [
        {
          "appId": "9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d",
          "name": "CLI app1"
        }
      ]
    }));
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      fs.existsSync,
      fs.readFileSync
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PERMISSION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves permissions from app registration if service principal not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/appRoles`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/appRoles`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.resolve({
            "value": [
              {
                "id": "5f348523-3353-4eba-8fe4-0af7a07eb872"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return Promise.resolve(appRegDelegatedPermissionsMultipleResources);
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '7df0a125-d3be-4c96-aa54-591f83ff541c'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "7df0a125-d3be-4c96-aa54-591f83ff541c",
                "id": "582d24e0-4dd7-41c5-b7dd-2a52817a95aa",
                "displayName": "Microsoft Flow Service"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
                "id": "c7c82441-65de-4fb1-ac2e-83a947ced55f",
                "displayName": "Windows Azure Service Management API"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/oauth2PermissionScopes`:
          return Promise.resolve(flowServiceOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/oauth2PermissionScopes`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves permissions from app registration if service principal not found (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/appRoles`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/appRoles`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.resolve({
            "value": [
              {
                "id": "5f348523-3353-4eba-8fe4-0af7a07eb872"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return Promise.resolve(appRegDelegatedPermissionsMultipleResources);
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '7df0a125-d3be-4c96-aa54-591f83ff541c'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "7df0a125-d3be-4c96-aa54-591f83ff541c",
                "id": "582d24e0-4dd7-41c5-b7dd-2a52817a95aa",
                "displayName": "Microsoft Flow Service"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
                "id": "c7c82441-65de-4fb1-ac2e-83a947ced55f",
                "displayName": "Windows Azure Service Management API"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/oauth2PermissionScopes`:
          return Promise.resolve(flowServiceOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/oauth2PermissionScopes`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves delegated permissions from app registration', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/appRoles`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/appRoles`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.resolve({
            "value": [
              {
                "id": "5f348523-3353-4eba-8fe4-0af7a07eb872"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return Promise.resolve(appRegDelegatedPermissionsMultipleResources);
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '7df0a125-d3be-4c96-aa54-591f83ff541c'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "7df0a125-d3be-4c96-aa54-591f83ff541c",
                "id": "582d24e0-4dd7-41c5-b7dd-2a52817a95aa",
                "displayName": "Microsoft Flow Service"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
                "id": "c7c82441-65de-4fb1-ac2e-83a947ced55f",
                "displayName": "Windows Azure Service Management API"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/oauth2PermissionScopes`:
          return Promise.resolve(flowServiceOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/oauth2PermissionScopes`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves application permissions from app registration', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.resolve({
            "value": [
              {
                "id": "5f348523-3353-4eba-8fe4-0af7a07eb872"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return Promise.resolve(appRegApplicationPermissions);
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
          {
            "resource": "Microsoft Graph",
            "permission": "AppCatalog.Read.All",
            "type": "Application"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`doesn't fail when the app registration has no API permissions`, (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.resolve({
            "value": [
              {
                "id": "5f348523-3353-4eba-8fe4-0af7a07eb872"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return Promise.resolve(appRegNoApiPermissions);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves permissions for a service principal with delegated and app permissions', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return Promise.resolve({
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves permissions for a service principal with delegated and app permissions (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return Promise.resolve({
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves permissions for a service principal with delegated permissions', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return Promise.resolve({ "value": [] });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return Promise.resolve({
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles non-existent app', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({ value: [] });
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('No Azure AD application registration with ID 9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving service principal for the AAD app', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.reject({
            error: {
              message: `An error has occurred`
            }
          });
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving OAuth2 permission grants for service principal', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return Promise.reject({
            error: {
              message: `An error has occurred`
            }
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return Promise.resolve({ value: [] });
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving OAuth2 permission scopes for service principal', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return Promise.resolve({
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.reject({
            error: {
              message: `An error has occurred`
            }
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving app role assignments for service principal', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return Promise.reject({
            error: {
              message: `An error has occurred`
            }
          });
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving app roles for service principal', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return Promise.resolve({
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.reject({
            error: {
              message: `An error has occurred`
            }
          });
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving AAD app registration', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.reject({
            error: {
              message: `An error has occurred`
            }
          });
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles non-existent service principal from app registration permissions', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.resolve({
            "value": [
              {
                "id": "5f348523-3353-4eba-8fe4-0af7a07eb872"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          return Promise.resolve(appRegApplicationPermissions);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
          {
            "resource": "00000003-0000-0000-c000-000000000000",
            "permission": "e12dae10-5a57-4817-b79d-dfbec5348930",
            "type": "Application"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles service principal referencing a non-existent app permission', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({
            value: [
              {
                "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
                "id": "14e36151-e472-4ece-812c-3e80a83fa3f5",
                "displayName": "CLI app"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/oauth2PermissionGrants`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/14e36151-e472-4ece-812c-3e80a83fa3f5/appRoleAssignments`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c?$select=appId,id,displayName`:
          return Promise.resolve({
            "appId": "00000003-0000-0000-c000-000000000000",
            "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
            "displayName": "Microsoft Graph"
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles unknown delegated permissions from app registration', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/appRoles`:
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/appRoles`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.resolve({
            "value": [
              {
                "id": "5f348523-3353-4eba-8fe4-0af7a07eb872"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          const appReg = appRegDelegatedPermissionsMultipleResources;
          appReg.requiredResourceAccess[0].resourceAccess[0].id = "e45c5562-459d-4d1b-8148-83eb1b6dcf84";
          return Promise.resolve(appReg);
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '7df0a125-d3be-4c96-aa54-591f83ff541c'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "7df0a125-d3be-4c96-aa54-591f83ff541c",
                "id": "582d24e0-4dd7-41c5-b7dd-2a52817a95aa",
                "displayName": "Microsoft Flow Service"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
                "id": "c7c82441-65de-4fb1-ac2e-83a947ced55f",
                "displayName": "Windows Azure Service Management API"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/582d24e0-4dd7-41c5-b7dd-2a52817a95aa/oauth2PermissionScopes`:
          return Promise.resolve(flowServiceOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/c7c82441-65de-4fb1-ac2e-83a947ced55f/oauth2PermissionScopes`:
          return Promise.resolve({
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
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles unknown application permissions from app registration', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      switch (opts.url) {
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=appId,id,displayName`:
          return Promise.resolve({ value: [] });
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d'&$select=id`:
          return Promise.resolve({
            "value": [
              {
                "id": "5f348523-3353-4eba-8fe4-0af7a07eb872"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/myorganization/applications/5f348523-3353-4eba-8fe4-0af7a07eb872`:
          const appReg = appRegApplicationPermissions;
          appReg.requiredResourceAccess[0].resourceAccess[0].id = 'e12dae10-5a57-4817-b79d-dfbec5348931';
          return Promise.resolve(appReg);
        case `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=appId,id,displayName`:
          return Promise.resolve({
            "value": [
              {
                "appId": "00000003-0000-0000-c000-000000000000",
                "id": "89d2b38d-6c40-46eb-b396-f6dfd70ff07c",
                "displayName": "Microsoft Graph"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/oauth2PermissionScopes`:
          return Promise.resolve(msGraphPrincipalOAuth2PermissionScopes);
        case `https://graph.microsoft.com/v1.0/servicePrincipals/89d2b38d-6c40-46eb-b396-f6dfd70ff07c/appRoles`:
          return Promise.resolve(msGraphPrincipalAppRoles);
        default:
          return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
      }
    });

    command.action(logger, { options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        assert.strictEqual(JSON.stringify(loggerLogSpy.lastCall.args[0]), JSON.stringify([
          {
            "resource": "Microsoft Graph",
            "permission": "e12dae10-5a57-4817-b79d-dfbec5348931",
            "type": "Application"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});