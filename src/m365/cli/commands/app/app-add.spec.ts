import assert from 'assert';
import fs from 'fs';
import forge from 'node-forge';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-add.js';
import * as mocks from './app-add.mock.js';
import { CommandError } from '../../../../Command.js';

describe(commands.APP_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoTenantId = '48526e9f-60c5-3000-31d7-aa1dc75ecf3c|908bel80-a04a-4422-b4a0-883d9847d110:c8e761e2-d528-34d1-8776-dc51157d619a&#xA;Tenant';
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      request.post,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync,
      forge.pki.rsa.generateKeyPair,
      forge.pki.createCertificate,
      forge.pki.certificateToPem,
      forge.pkcs12.toPkcs12Asn1,
      forge.asn1.toDer,
      forge.util.encode64
    ]);
    (command as any).manifest = undefined;
  });

  after(() => {
    sinon.restore();
    //auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if name, delegated mode and permissions are defined', async () => {
    const actual = await command.validate({ options: { name: 'My Microsoft Entra app', mode: 'delegated', permissions: 'https://graph.microsoft.com/Group.ReadWrite.All' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if name, appOnly mode and permissions are defined', async () => {
    const actual = await command.validate({ options: { name: 'My Microsoft Entra app', mode: 'appOnly', permissions: 'https://graph.microsoft.com/Group.ReadWrite.All' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if name, delegated mode and correct permissionSet are defined', async () => {
    const actual = await command.validate({ options: { name: 'My Microsoft Entra app', mode: 'delegated', permissionSet: 'SpoRead' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if name, appOnly mode and correct permissionSet are defined', async () => {
    const actual = await command.validate({ options: { name: 'My Microsoft Entra app', mode: 'appOnly', permissionSet: 'SpoRead' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if permissionSet is incorrect', async () => {
    const actual = await command.validate({ options: { name: 'My Microsoft Entra app', mode: 'appOnly', permissionSet: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if permissionSet and permissions are defined together', async () => {
    const actual = await command.validate({ options: { name: 'My Microsoft Entra app', mode: 'appOnly', permissionSet: 'foo', permissions: 'https://graph.microsoft.com/Group.ReadWrite.All' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if permissionSet and permissions are not defined', async () => {
    const actual = await command.validate({ options: { name: 'My Microsoft Entra app', mode: 'appOnly' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if mode is incorrect', async () => {
    const actual = await command.validate({ options: { name: 'My Microsoft Entra app', mode: 'foo', permissionSet: 'SpoFull' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('creates correct Microsoft Entra app reg in delegated mode with defined permission', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "4e46008b-f24c-477d-8fff-7bb4ec7aafe0",
                  "type": "Scope"
                }
              ]
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    await command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'delegated', permissions: 'https://graph.microsoft.com/Group.ReadWrite.All', debug: true
      }
    });

    assert(loggerLogSpy.calledWith({
      appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
      name: 'My Microsoft Entra app',
      objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      tenantId: ''
    }));
  });

  it('creates correct Microsoft Entra app reg in delegated mode with SpoRead permissionSet', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "4e0d77b0-96ba-4398-af14-3baa780278f4",
                  "type": "Scope"
                },
                {
                  "id": "a468ea40-458c-4cc2-80c4-51781af71e41",
                  "type": "Scope"
                },
                {
                  "id": "0cea5a30-f6f8-42b5-87a0-84cc26822e02",
                  "type": "Scope"
                }
              ]
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    await command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'delegated', permissionSet: 'SpoRead', debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
      name: 'My Microsoft Entra app',
      objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      tenantId: ''
    }));
  });

  it('creates correct Microsoft Entra app reg in delegated mode with SpoFull permissionSet', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "56680e0d-d2a3-4ae1-80d8-3c4f2100e3d0",
                  "type": "Scope"
                },
                {
                  "id": "59a198b5-0420-45a8-ae59-6da1cb640505",
                  "type": "Scope"
                },
                {
                  "id": "0cea5a30-f6f8-42b5-87a0-84cc26822e02",
                  "type": "Scope"
                }
              ]
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    await command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'delegated', permissionSet: 'SpoFull', debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
      name: 'My Microsoft Entra app',
      objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      tenantId: ''
    }));
  });

  it('creates correct Microsoft Entra app reg in delegated mode with Full permissionSet', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult,
            mocks.yammerSP,
            mocks.managementAzureSp,
            mocks.managementOfficeSp
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
              "resourceAccess": [
                {
                  "id": "41094075-9dad-400e-a0bd-54e686782033",
                  "type": "Scope"
                }
              ]
            },
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "1ca167d5-1655-44a1-8adf-1414072e1ef9",
                  "type": "Scope"
                },
                {
                  "id": "0e263e50-5827-48a4-b97c-d940288653c7",
                  "type": "Scope"
                },
                {
                  "id": "c5366453-9fb0-48a5-a156-24f0c49a4b84",
                  "type": "Scope"
                },
                {
                  "id": "4e46008b-f24c-477d-8fff-7bb4ec7aafe0",
                  "type": "Scope"
                },
                {
                  "id": "f13ce604-1677-429f-90bd-8a10b9f01325",
                  "type": "Scope"
                },
                {
                  "id": "024d486e-b451-40bb-833d-3e66d98c5c73",
                  "type": "Scope"
                },
                {
                  "id": "e383f46e-2787-4529-855e-0e479a3ffac0",
                  "type": "Scope"
                },
                {
                  "id": "02e97553-ed7b-43d0-ab3c-f8bace0d040c",
                  "type": "Scope"
                },
                {
                  "id": "2219042f-cab5-40cc-b0d2-16b1540b4c5f",
                  "type": "Scope"
                },
                {
                  "id": "093f8818-d05f-49b8-95bc-9d2a73e9a43c",
                  "type": "Scope"
                },
                {
                  "id": "63dd7cd9-b489-4adf-a28c-ac38b9a0f962",
                  "type": "Scope"
                }
              ]
            },
            {
              "resourceAppId": "c5393580-f805-4401-95e8-94b7a6ef2fc2",
              "resourceAccess": [
                {
                  "id": "e2cea78f-e743-4d8f-a16a-75b629a038ae",
                  "type": "Scope"
                }
              ]
            },
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "56680e0d-d2a3-4ae1-80d8-3c4f2100e3d0",
                  "type": "Scope"
                },
                {
                  "id": "59a198b5-0420-45a8-ae59-6da1cb640505",
                  "type": "Scope"
                },
                {
                  "id": "0cea5a30-f6f8-42b5-87a0-84cc26822e02",
                  "type": "Scope"
                }
              ]
            },
            {
              "resourceAppId": "00000005-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "5db81a03-0de0-432b-b31e-71d57c8d2e0b",
                  "type": "Scope"
                }
              ]
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    await command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'delegated', permissionSet: true, debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
      name: 'My Microsoft Entra app',
      objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      tenantId: ''
    }));
  });

  it('creates correct Microsoft Entra app reg in delegated mode with ReadAll permissionSet', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult,
            mocks.yammerSP,
            mocks.managementAzureSp,
            mocks.managementOfficeSp
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "797f4846-ba00-4fd7-ba43-dac1f8f63013",
              "resourceAccess": [
                {
                  "id": "41094075-9dad-400e-a0bd-54e686782033",
                  "type": "Scope"
                }
              ]
            },
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "88e58d74-d3df-44f3-ad47-e89edf4472e4",
                  "type": "Scope"
                },
                {
                  "id": "0e263e50-5827-48a4-b97c-d940288653c7",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                },
                {
                  "id": "5f8c59db-677d-491f-a6b8-5f174b11ec1d",
                  "type": "Scope"
                },
                {
                  "id": "43781733-b5a7-4d1b-98f4-e8edff23e1a9",
                  "type": "Scope"
                },
                {
                  "id": "570282fd-fa5c-430d-a7fd-fc8dc98a9dca",
                  "type": "Scope"
                },
                {
                  "id": "02e97553-ed7b-43d0-ab3c-f8bace0d040c",
                  "type": "Scope"
                },
                {
                  "id": "f45671fb-e0fe-4b4b-be20-3d3ce43f1bcb",
                  "type": "Scope"
                },
                {
                  "id": "c395395c-ff9a-4dba-bc1f-8372ba9dca84",
                  "type": "Scope"
                }
              ]
            },
            {
              "resourceAppId": "c5393580-f805-4401-95e8-94b7a6ef2fc2",
              "resourceAccess": [
                {
                  "id": "e2cea78f-e743-4d8f-a16a-75b629a038ae",
                  "type": "Scope"
                }
              ]
            },
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "4e0d77b0-96ba-4398-af14-3baa780278f4",
                  "type": "Scope"
                },
                {
                  "id": "a468ea40-458c-4cc2-80c4-51781af71e41",
                  "type": "Scope"
                },
                {
                  "id": "0cea5a30-f6f8-42b5-87a0-84cc26822e02",
                  "type": "Scope"
                }
              ]
            },
            {
              "resourceAppId": "00000005-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "5db81a03-0de0-432b-b31e-71d57c8d2e0b",
                  "type": "Scope"
                }
              ]
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    await command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'delegated', permissionSet: 'ReadAll', debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
      name: 'My Microsoft Entra app',
      objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      tenantId: ''
    }));
  });

  it('creates correct Microsoft Entra app reg in appOnly mode with defined permission', async () => {
    const certificateEncoded64Result = 'certificate-result';
    let passwordResult = 'incorrect-password';
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "62a82d76-70ea-41e2-9197-370581804d09",
                  "type": "Role"
                }
              ]
            }
          ],
          "keyCredentials": [
            {
              "type": 'AsymmetricX509Cert',
              "usage": 'Verify',
              "displayName": 'PnP M365 Management Shell',
              "key": certificateEncoded64Result
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [
            {
              "customKeyIdentifier": "customKeyIdentifier",
              "endDateTime": "endDateTime"
            }
          ],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/59e617e5-e447-4adc-8b88-00af644d7c92/appRoleAssignments') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(forge.pki.rsa, 'generateKeyPair').callsFake(() => {
      return {
        publicKey: 'test-key',
        privateKey: 'test-key'
      } as any;
    });

    sinon.stub(forge.pki, 'createCertificate').callsFake(() => {
      return {
        validity: {},
        setSubject: () => { },
        setIssuer: () => { },
        sign: () => { }
      } as any;
    });
    sinon.stub(forge.pki, 'certificateToPem').callsFake(() => 'test-key');
    sinon.stub(forge.pkcs12, 'toPkcs12Asn1').callsFake((privateKey: any, certificate: any, password: any) => {
      privateKey = privateKey;
      certificate = certificate;
      passwordResult = password;

      return {} as any;
    });
    sinon.stub(forge.asn1, 'toDer').callsFake(() => {
      return {
        getBytes: () => { }
      } as any;
    });
    sinon.stub(forge.util, 'encode64').callsFake(() => certificateEncoded64Result);
    sinon.stub(String.prototype, 'charAt').callsFake(() => 'A');

    await command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'appOnly', permissions: 'https://graph.microsoft.com/Group.ReadWrite.All', debug: true
      }
    });

    assert(loggerLogSpy.calledWith({
      appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
      objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      tenantId: '',
      name: 'My Microsoft Entra app',
      certPassword: passwordResult,
      certThumbprint: 'customKeyIdentifier',
      certExpirationDate: 'endDateTime'
    }));
  });

  it('creates correct Microsoft Entra app reg in appOnly mode with chosen permissionSet', async () => {
    const certificateEncoded64Result = 'certificate-result';
    let passwordResult = 'incorrect-password';
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "d13f72ca-a275-4b96-b789-48ebcc4da984",
                  "type": "Role"
                },
                {
                  "id": "2a8d57a5-4090-4a41-bf1c-3c621d2ccad3",
                  "type": "Role"
                },
                {
                  "id": "df021288-bdef-4463-88db-98f22de89214",
                  "type": "Role"
                }
              ]
            }
          ],
          keyCredentials: [
            {
              type: 'AsymmetricX509Cert',
              usage: 'Verify',
              displayName: 'PnP M365 Management Shell',
              key: 'certificate-result'
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [
            {
              "customKeyIdentifier": "customKeyIdentifier",
              "endDateTime": "endDateTime"
            }
          ],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/59e617e5-e447-4adc-8b88-00af644d7c92/appRoleAssignments') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(forge.pki.rsa, 'generateKeyPair').callsFake(() => {
      return {
        publicKey: 'test-key',
        privateKey: 'test-key'
      } as any;
    });

    sinon.stub(forge.pki, 'createCertificate').callsFake(() => {
      return {
        validity: {},
        setSubject: () => { },
        setIssuer: () => { },
        sign: () => { }
      } as any;
    });
    sinon.stub(forge.pki, 'certificateToPem').callsFake(() => 'test-key');
    sinon.stub(forge.pkcs12, 'toPkcs12Asn1').callsFake((privateKey: any, certificate: any, password: any) => {
      privateKey = privateKey;
      certificate = certificate;
      passwordResult = password;

      return {} as any;
    });
    sinon.stub(forge.asn1, 'toDer').callsFake(() => {
      return {
        getBytes: () => { }
      } as any;
    });
    sinon.stub(forge.util, 'encode64').callsFake(() => certificateEncoded64Result);

    await command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'appOnly', permissionSet: 'SpoRead', debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
      objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      tenantId: '',
      name: 'My Microsoft Entra app',
      certPassword: passwordResult,
      certThumbprint: 'customKeyIdentifier',
      certExpirationDate: 'endDateTime'
    }));
  });

  it('creates correct Microsoft Entra app reg in appOnly mode with chosen permissionSet when in response keyCredentials property is empty', async () => {
    const certificateEncoded64Result = 'certificate-result';
    let passwordResult = 'incorrect-password';
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "d13f72ca-a275-4b96-b789-48ebcc4da984",
                  "type": "Role"
                },
                {
                  "id": "2a8d57a5-4090-4a41-bf1c-3c621d2ccad3",
                  "type": "Role"
                },
                {
                  "id": "df021288-bdef-4463-88db-98f22de89214",
                  "type": "Role"
                }
              ]
            }
          ],
          keyCredentials: [
            {
              type: 'AsymmetricX509Cert',
              usage: 'Verify',
              displayName: 'PnP M365 Management Shell',
              key: 'certificate-result'
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [
          ],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/59e617e5-e447-4adc-8b88-00af644d7c92/appRoleAssignments') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(forge.pki.rsa, 'generateKeyPair').callsFake(() => {
      return {
        publicKey: 'test-key',
        privateKey: 'test-key'
      } as any;
    });

    sinon.stub(forge.pki, 'createCertificate').callsFake(() => {
      return {
        validity: {},
        setSubject: () => { },
        setIssuer: () => { },
        sign: () => { }
      } as any;
    });
    sinon.stub(forge.pki, 'certificateToPem').callsFake(() => 'test-key');
    sinon.stub(forge.pkcs12, 'toPkcs12Asn1').callsFake((privateKey: any, certificate: any, password: any) => {
      privateKey = privateKey;
      certificate = certificate;
      passwordResult = password;

      return {} as any;
    });
    sinon.stub(forge.asn1, 'toDer').callsFake(() => {
      return {
        getBytes: () => { }
      } as any;
    });
    sinon.stub(forge.util, 'encode64').callsFake(() => certificateEncoded64Result);

    await command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'appOnly', permissionSet: 'SpoRead', debug: true
      }
    });
    assert(loggerLogSpy.calledWith({
      appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
      objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
      tenantId: '',
      name: 'My Microsoft Entra app',
      certPassword: passwordResult,
      certThumbprint: undefined,
      certExpirationDate: undefined
    }));
  });

  it('throws error when non-existent permission is specified permission option', async () => {
    const certificateEncoded64Result = 'certificate-result';
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "df021288-bdef-4463-88db-98f22de89214",
                  "type": "Role"
                },
                {
                  "id": "2a8d57a5-4090-4a41-bf1c-3c621d2ccad3",
                  "type": "Role"
                },
                {
                  "id": "d13f72ca-a275-4b96-b789-48ebcc4da984",
                  "type": "Role"
                }
              ]
            }
          ],
          keyCredentials: [
            {
              type: 'AsymmetricX509Cert',
              usage: 'Verify',
              displayName: 'PnP M365 Management Shell',
              key: 'certificate-result'
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [
            {
              "customKeyIdentifier": "customKeyIdentifier",
              "endDateTime": "endDateTime"
            }
          ],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/59e617e5-e447-4adc-8b88-00af644d7c92/appRoleAssignments') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(forge.pki.rsa, 'generateKeyPair').callsFake(() => {
      return {
        publicKey: 'test-key',
        privateKey: 'test-key'
      } as any;
    });

    sinon.stub(forge.pki, 'createCertificate').callsFake(() => {
      return {
        validity: {},
        setSubject: () => { },
        setIssuer: () => { },
        sign: () => { }
      } as any;
    });
    sinon.stub(forge.pki, 'certificateToPem').callsFake(() => 'test-key');
    sinon.stub(forge.pkcs12, 'toPkcs12Asn1').callsFake(() => {
      return {} as any;
    });
    sinon.stub(forge.asn1, 'toDer').callsFake(() => {
      return {
        getBytes: () => { }
      } as any;
    });
    sinon.stub(forge.util, 'encode64').callsFake(() => certificateEncoded64Result);

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'appOnly', permissions: 'https://graph.microsoft.com/Read.Everything'
      }
    } as any), new CommandError('Permission Read.Everything for service principal https://graph.microsoft.com not found'));
  });

  it('throws error when non-existent service principal is specified in the permissions option', async () => {

    const certificateEncoded64Result = 'certificate-result';
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "df021288-bdef-4463-88db-98f22de89214",
                  "type": "Role"
                },
                {
                  "id": "2a8d57a5-4090-4a41-bf1c-3c621d2ccad3",
                  "type": "Role"
                },
                {
                  "id": "d13f72ca-a275-4b96-b789-48ebcc4da984",
                  "type": "Role"
                }
              ]
            }
          ],
          keyCredentials: [
            {
              type: 'AsymmetricX509Cert',
              usage: 'Verify',
              displayName: 'PnP M365 Management Shell',
              key: 'certificate-result'
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [
            {
              "customKeyIdentifier": "customKeyIdentifier",
              "endDateTime": "endDateTime"
            }
          ],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/59e617e5-e447-4adc-8b88-00af644d7c92/appRoleAssignments') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(forge.pki.rsa, 'generateKeyPair').callsFake(() => {
      return {
        publicKey: 'test-key',
        privateKey: 'test-key'
      } as any;
    });

    sinon.stub(forge.pki, 'createCertificate').callsFake(() => {
      return {
        validity: {},
        setSubject: () => { },
        setIssuer: () => { },
        sign: () => { }
      } as any;
    });
    sinon.stub(forge.pki, 'certificateToPem').callsFake(() => 'test-key');
    sinon.stub(forge.pkcs12, 'toPkcs12Asn1').callsFake(() => {
      return {} as any;
    });
    sinon.stub(forge.asn1, 'toDer').callsFake(() => {
      return {
        getBytes: () => { }
      } as any;
    });
    sinon.stub(forge.util, 'encode64').callsFake(() => certificateEncoded64Result);

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'appOnly', permissions: 'https://myapi.onmicrosoft.com/access_as_user'
      }
    }), new CommandError('Service principal https://myapi.onmicrosoft.com not found'));
  });

  it('throws error when certificate process fails', async () => {

    const certificateEncoded64Result = 'certificate-result';
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames') {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.graphResult
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId%2cappRoles%2cid%2coauth2PermissionScopes%2cservicePrincipalNames&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return {
          value: mocks.sharePointSp
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });
    sinon.stub(request, 'patch').rejects('Issued PATCH request');
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          },
          "isFallbackPublicClient": true,
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0ff1-ce00-000000000000",
              "resourceAccess": [
                {
                  "id": "df021288-bdef-4463-88db-98f22de89214",
                  "type": "Role"
                },
                {
                  "id": "2a8d57a5-4090-4a41-bf1c-3c621d2ccad3",
                  "type": "Role"
                },
                {
                  "id": "d13f72ca-a275-4b96-b789-48ebcc4da984",
                  "type": "Role"
                }
              ]
            }
          ],
          keyCredentials: [
            {
              type: 'AsymmetricX509Cert',
              usage: 'Verify',
              displayName: 'PnP M365 Management Shell',
              key: 'certificate-result'
            }
          ]
        })) {
        return {
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My Microsoft Entra app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": []
          },
          "defaultRedirectUri": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [],
            "preAuthorizedApplications": []
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [
            {
              "customKeyIdentifier": "customKeyIdentifier",
              "endDateTime": "endDateTime"
            }
          ],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals') {
        return {
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "displayName": "My Microsoft Entra app",
          "appRoles": [],
          "oauth2PermissionScopes": [],
          "servicePrincipalNames": [
            "f1bd758f-4a1a-4b71-aa20-a248a22a8928"
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/59e617e5-e447-4adc-8b88-00af644d7c92/appRoleAssignments') {
        return;
      }

      throw `Invalid POST request: ${JSON.stringify(opts, null, 2)}`;
    });

    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(forge.pki.rsa, 'generateKeyPair').callsFake(() => {
      return {
        publicKey: 'test-key',
        privateKey: 'test-key'
      } as any;
    });

    sinon.stub(forge.pki, 'createCertificate').callsFake(() => {
      throw 'Certificate creation error';
    });
    sinon.stub(forge.pki, 'certificateToPem').callsFake(() => 'test-key');
    sinon.stub(forge.pkcs12, 'toPkcs12Asn1').callsFake(() => {
      return {} as any;
    });
    sinon.stub(forge.asn1, 'toDer').callsFake(() => {
      return {
        getBytes: () => { }
      } as any;
    });
    sinon.stub(forge.util, 'encode64').callsFake(() => certificateEncoded64Result);

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My Microsoft Entra app', mode: 'appOnly', permissions: 'https://graph.microsoft.com/Group.ReadWrite.All'
      }
    }), new CommandError('Error while creating certificate file: Certificate creation error.'));
  });
});
