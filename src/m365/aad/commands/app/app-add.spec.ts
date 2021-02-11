import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import * as mocks from './app-add.mock';
const command: Command = require('./app-add');

describe(commands.APP_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.tenantId = '48526e9f-60c5-3000-31d7-aa1dc75ecf3c|908bel80-a04a-4422-b4a0-883d9847d110:c8e761e2-d528-34d1-8776-dc51157d619a&#xA;Tenant';
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
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
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.patch,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates AAD app reg with just the name', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issues GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
          "deletedDateTime": null,
          "appId": "bc724b77-da87-43a9-b385-6ebaaf969db8",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:44:13.7945807Z",
          "displayName": "My AAD app",
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
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
          objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates multitenant AAD app reg', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issues GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMultipleOrgs"
        })) {
        return Promise.resolve({
          "id": "9b1e2c08-6e35-4134-a0ac-16ab154cd05a",
          "deletedDateTime": null,
          "appId": "62f0f128-987f-47f2-827a-be50d0d894c7",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:50:40.1806422Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
          "signInAudience": "AzureADMultipleOrgs",
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
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        multitenant: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: '62f0f128-987f-47f2-827a-be50d0d894c7',
          objectId: '9b1e2c08-6e35-4134-a0ac-16ab154cd05a',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg for a web app with the specified redirect URIs', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issues GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "web": {
            "redirectUris": [
              "https://myapp.azurewebsites.net",
              "http://localhost:4000"
            ]
          }
        })) {
        return Promise.resolve({
          "id": "ff520671-4810-4d25-a10f-e565fc62a5ec",
          "deletedDateTime": null,
          "appId": "d2941a3b-aad4-49e0-8a1d-b82de0b46067",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:53:40.7071625Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
            "redirectUris": [
              "https://myapp.azurewebsites.net",
              "http://localhost:4000"
            ],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        redirectUris: 'https://myapp.azurewebsites.net,http://localhost:4000',
        platform: 'web'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: 'd2941a3b-aad4-49e0-8a1d-b82de0b46067',
          objectId: 'ff520671-4810-4d25-a10f-e565fc62a5ec',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg for a desktop app with the specified redirect URI', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issues GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "publicClient": {
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
          }
        })) {
        return Promise.resolve({
          "id": "f1bb2138-bff1-491e-b082-9f447f3742b8",
          "deletedDateTime": null,
          "appId": "1ce0287c-9ccc-457e-a0cf-3ec5b734c092",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:56:17.4207858Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
            "redirectUris": [
              "https://login.microsoftonline.com/common/oauth2/nativeclient"
            ]
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
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        redirectUris: 'https://login.microsoftonline.com/common/oauth2/nativeclient',
        platform: 'publicClient'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: '1ce0287c-9ccc-457e-a0cf-3ec5b734c092',
          objectId: 'f1bb2138-bff1-491e-b082-9f447f3742b8',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg with a secret', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issues GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "4d24b0c6-ad07-47c6-9bd8-9c167f9f758e",
          "deletedDateTime": null,
          "appId": "3c5bd51d-f1ac-4344-bd16-43396cadff14",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:58:18.7120335Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/4d24b0c6-ad07-47c6-9bd8-9c167f9f758e/addPassword') {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#microsoft.graph.passwordCredential",
          "customKeyIdentifier": null,
          "displayName": "Default",
          "endDateTime": "2120-12-31T14:58:16.875Z",
          "hint": "VtJ",
          "keyId": "17dc40d4-7c81-47dd-a3cb-41df4aed1130",
          "secretText": "VtJt.yG~V5pzbY2.xekx_0Xy_~9ozP_Ub5",
          "startDateTime": "2020-12-31T14:58:19.2307535Z"
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        withSecret: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: '3c5bd51d-f1ac-4344-bd16-43396cadff14',
          objectId: '4d24b0c6-ad07-47c6-9bd8-9c167f9f758e',
          secret: 'VtJt.yG~V5pzbY2.xekx_0Xy_~9ozP_Ub5',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg with a secret (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issues GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "4d24b0c6-ad07-47c6-9bd8-9c167f9f758e",
          "deletedDateTime": null,
          "appId": "3c5bd51d-f1ac-4344-bd16-43396cadff14",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:58:18.7120335Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/4d24b0c6-ad07-47c6-9bd8-9c167f9f758e/addPassword') {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#microsoft.graph.passwordCredential",
          "customKeyIdentifier": null,
          "displayName": "Default",
          "endDateTime": "2120-12-31T14:58:16.875Z",
          "hint": "VtJ",
          "keyId": "17dc40d4-7c81-47dd-a3cb-41df4aed1130",
          "secretText": "VtJt.yG~V5pzbY2.xekx_0Xy_~9ozP_Ub5",
          "startDateTime": "2020-12-31T14:58:19.2307535Z"
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'My AAD app',
        withSecret: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: '3c5bd51d-f1ac-4344-bd16-43396cadff14',
          objectId: '4d24b0c6-ad07-47c6-9bd8-9c167f9f758e',
          secret: 'VtJt.yG~V5pzbY2.xekx_0Xy_~9ozP_Ub5',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg for a deamon app with specified Microsoft Graph application permissions', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles') {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.microsoftGraphSp
          ]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return Promise.resolve({
          value: mocks.aadSp
        });
      }

      return Promise.reject(`Invalid GET request: ${opts.url}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "62a82d76-70ea-41e2-9197-370581804d09",
                  "type": "Role"
                },
                {
                  "id": "7ab1d382-f21e-4acd-a863-ba3e13f7da61",
                  "type": "Role"
                }
              ]
            }
          ]
        })) {
        return Promise.resolve({
          "id": "b63c4be1-9c78-40b7-8619-de7172eed8de",
          "deletedDateTime": null,
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T15:02:42.8048505Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "62a82d76-70ea-41e2-9197-370581804d09",
                  "type": "Role"
                },
                {
                  "id": "7ab1d382-f21e-4acd-a863-ba3e13f7da61",
                  "type": "Role"
                }
              ]
            }
          ],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/b63c4be1-9c78-40b7-8619-de7172eed8de/addPassword') {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#microsoft.graph.passwordCredential",
          "customKeyIdentifier": null,
          "displayName": "Default",
          "endDateTime": "2120-12-31T15:02:40.978Z",
          "hint": "vP2",
          "keyId": "f7394450-52f6-4c04-926c-dc29398eaa1c",
          "secretText": "vP2K-_K-N6EI-E5z0yOTsz443grfM_pyvv",
          "startDateTime": "2020-12-31T15:02:43.2435402Z"
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        withSecret: true,
        apisApplication: 'https://graph.microsoft.com/Group.ReadWrite.All,https://graph.microsoft.com/Directory.Read.All'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: 'dbfdad7a-5105-45fc-8290-eb0b0b24ac58',
          objectId: 'b63c4be1-9c78-40b7-8619-de7172eed8de',
          secret: 'vP2K-_K-N6EI-E5z0yOTsz443grfM_pyvv',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg for a deamon app with specified Microsoft Graph application and delegated permissions', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles') {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.microsoftGraphSp
          ]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return Promise.resolve({
          value: mocks.aadSp
        });
      }

      return Promise.reject(`Invalid GET request: ${opts.url}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                },
                {
                  "id": "62a82d76-70ea-41e2-9197-370581804d09",
                  "type": "Role"
                },
                {
                  "id": "7ab1d382-f21e-4acd-a863-ba3e13f7da61",
                  "type": "Role"
                }
              ]
            }
          ]
        })) {
        return Promise.resolve({
          "id": "b63c4be1-9c78-40b7-8619-de7172eed8de",
          "deletedDateTime": null,
          "appId": "dbfdad7a-5105-45fc-8290-eb0b0b24ac58",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T15:02:42.8048505Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                },
                {
                  "id": "62a82d76-70ea-41e2-9197-370581804d09",
                  "type": "Role"
                },
                {
                  "id": "7ab1d382-f21e-4acd-a863-ba3e13f7da61",
                  "type": "Role"
                }
              ]
            }
          ],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          }
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/b63c4be1-9c78-40b7-8619-de7172eed8de/addPassword') {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#microsoft.graph.passwordCredential",
          "customKeyIdentifier": null,
          "displayName": "Default",
          "endDateTime": "2120-12-31T15:02:40.978Z",
          "hint": "vP2",
          "keyId": "f7394450-52f6-4c04-926c-dc29398eaa1c",
          "secretText": "vP2K-_K-N6EI-E5z0yOTsz443grfM_pyvv",
          "startDateTime": "2020-12-31T15:02:43.2435402Z"
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        withSecret: true,
        apisApplication: 'https://graph.microsoft.com/Group.ReadWrite.All,https://graph.microsoft.com/Directory.Read.All',
        apisDelegated: 'https://graph.microsoft.com/Directory.Read.All'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', `Error: ${JSON.stringify(err, null, 2)}`);
        assert(loggerLogSpy.calledWith({
          appId: 'dbfdad7a-5105-45fc-8290-eb0b0b24ac58',
          objectId: 'b63c4be1-9c78-40b7-8619-de7172eed8de',
          secret: 'vP2K-_K-N6EI-E5z0yOTsz443grfM_pyvv',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg for a single-page app with specified Microsoft Graph delegated permissions', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles') {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.microsoftGraphSp
          ]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return Promise.resolve({
          value: mocks.aadSp
        });
      }

      return Promise.reject(`Invalid GET request: ${opts.url}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                }
              ]
            }
          ],
          "spa": {
            "redirectUris": [
              "https://myspa.azurewebsites.net",
              "http://localhost:8080"
            ]
          },
          "web": {
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": true,
              "enableIdTokenIssuance": true
            }
          }
        })) {
        return Promise.resolve({
          "id": "f51ff52f-8f04-4924-91d0-636349eed65c",
          "deletedDateTime": null,
          "appId": "c505d465-9e4e-4bb4-b653-7b36d77cc94a",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:08:27.9188248Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": [
              "https://myspa.azurewebsites.net",
              "http://localhost:8080"
            ]
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
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                }
              ]
            }
          ],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": true,
              "enableIdTokenIssuance": true
            }
          }
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        platform: 'spa',
        redirectUris: 'https://myspa.azurewebsites.net,http://localhost:8080',
        apisDelegated: 'https://graph.microsoft.com/Calendars.Read,https://graph.microsoft.com/Directory.Read.All',
        implicitFlow: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: 'c505d465-9e4e-4bb4-b653-7b36d77cc94a',
          objectId: 'f51ff52f-8f04-4924-91d0-636349eed65c',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg for a single-page app with specified Microsoft Graph delegated permissions (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles') {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.microsoftGraphSp
          ]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return Promise.resolve({
          value: mocks.aadSp
        });
      }

      return Promise.reject(`Invalid GET request: ${opts.url}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                }
              ]
            }
          ],
          "spa": {
            "redirectUris": [
              "https://myspa.azurewebsites.net",
              "http://localhost:8080"
            ]
          },
          "web": {
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": true,
              "enableIdTokenIssuance": true
            }
          }
        })) {
        return Promise.resolve({
          "id": "f51ff52f-8f04-4924-91d0-636349eed65c",
          "deletedDateTime": null,
          "appId": "c505d465-9e4e-4bb4-b653-7b36d77cc94a",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:08:27.9188248Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": [
              "https://myspa.azurewebsites.net",
              "http://localhost:8080"
            ]
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
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                }
              ]
            }
          ],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": true,
              "enableIdTokenIssuance": true
            }
          }
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'My AAD app',
        platform: 'spa',
        redirectUris: 'https://myspa.azurewebsites.net,http://localhost:8080',
        apisDelegated: 'https://graph.microsoft.com/Calendars.Read,https://graph.microsoft.com/Directory.Read.All',
        implicitFlow: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: 'c505d465-9e4e-4bb4-b653-7b36d77cc94a',
          objectId: 'f51ff52f-8f04-4924-91d0-636349eed65c',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg with Application ID URI set to a fixed value', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issued GET request'));
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/c0e63919-057c-4e6b-be6c-8662e7aec4eb' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "identifierUris": [
            "https://contoso.onmicrosoft.com/myapp"
          ]
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid PATCH request: ${JSON.stringify(opts, null, 2)}`);
    });
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "c0e63919-057c-4e6b-be6c-8662e7aec4eb",
          "deletedDateTime": null,
          "appId": "b08d9318-5612-4f87-9f94-7414ef6f0c8a",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:14:23.9641082Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        uri: 'https://contoso.onmicrosoft.com/myapp'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: 'b08d9318-5612-4f87-9f94-7414ef6f0c8a',
          objectId: 'c0e63919-057c-4e6b-be6c-8662e7aec4eb',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg with Application ID URI set to a fixed value (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issued GET request'));
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/c0e63919-057c-4e6b-be6c-8662e7aec4eb' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "identifierUris": [
            "https://contoso.onmicrosoft.com/myapp"
          ]
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid PATCH request: ${JSON.stringify(opts, null, 2)}`);
    });
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "c0e63919-057c-4e6b-be6c-8662e7aec4eb",
          "deletedDateTime": null,
          "appId": "b08d9318-5612-4f87-9f94-7414ef6f0c8a",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:14:23.9641082Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'My AAD app',
        uri: 'https://contoso.onmicrosoft.com/myapp'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: 'b08d9318-5612-4f87-9f94-7414ef6f0c8a',
          objectId: 'c0e63919-057c-4e6b-be6c-8662e7aec4eb',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg with Application ID URI set to a value with the appId token and a custom scope that can be consented by admins', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issued GET request'));
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/fe45ba27-a692-4b11-adf8-f4ec184ea3a5') {
        const actualData = JSON.stringify(opts.data);
        const expectedData = JSON.stringify({
          "identifierUris": [
            "api://caf406b91cd4.ngrok.io/13e11551-2967-4985-8c55-cd2aaa6b80ad"
          ],
          "api": {
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "|",
                "type": "Admin",
                "value": "access_as_user"
              }
            ]
          }
        }).split('|');
        if (actualData.indexOf(expectedData[0]) > -1 && actualData.indexOf(expectedData[1]) > -1) {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid PATCH request: ${JSON.stringify(opts, null, 2)}`);
    });
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "fe45ba27-a692-4b11-adf8-f4ec184ea3a5",
          "deletedDateTime": null,
          "appId": "13e11551-2967-4985-8c55-cd2aaa6b80ad",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:17:55.8423122Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        uri: 'api://caf406b91cd4.ngrok.io/_appId_',
        scopeName: 'access_as_user',
        scopeAdminConsentDescription: 'Access as a user',
        scopeAdminConsentDisplayName: 'Access as a user',
        scopeConsentBy: 'admins'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', `Error: ${JSON.stringify(err)}`);
        assert(loggerLogSpy.calledWith({
          appId: '13e11551-2967-4985-8c55-cd2aaa6b80ad',
          objectId: 'fe45ba27-a692-4b11-adf8-f4ec184ea3a5',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg with Application ID URI set to a value with the appId token and a custom scope that can be consented by admins and users', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issued GET request'));
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/fe45ba27-a692-4b11-adf8-f4ec184ea3a5') {
        const actualData = JSON.stringify(opts.data);
        const expectedData = JSON.stringify({
          "identifierUris": [
            "api://caf406b91cd4.ngrok.io/13e11551-2967-4985-8c55-cd2aaa6b80ad"
          ],
          "api": {
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "|",
                "type": "User",
                "value": "access_as_user"
              }
            ]
          }
        }).split('|');
        if (actualData.indexOf(expectedData[0]) > -1 && actualData.indexOf(expectedData[1]) > -1) {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid PATCH request: ${JSON.stringify(opts, null, 2)}`);
    });
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "fe45ba27-a692-4b11-adf8-f4ec184ea3a5",
          "deletedDateTime": null,
          "appId": "13e11551-2967-4985-8c55-cd2aaa6b80ad",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:17:55.8423122Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        uri: 'api://caf406b91cd4.ngrok.io/_appId_',
        scopeName: 'access_as_user',
        scopeAdminConsentDescription: 'Access as a user',
        scopeAdminConsentDisplayName: 'Access as a user',
        scopeConsentBy: 'adminsAndUsers'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', `Error: ${JSON.stringify(err)}`);
        assert(loggerLogSpy.calledWith({
          appId: '13e11551-2967-4985-8c55-cd2aaa6b80ad',
          objectId: 'fe45ba27-a692-4b11-adf8-f4ec184ea3a5',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when retrieving information about service principals failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject({
      error: {
        message: `An error has occurred`
      }
    }));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(_ => Promise.reject('Issued POST request'));

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        withSecret: true,
        apisApplication: 'https://graph.microsoft.com/Group.ReadWrite.All,https://graph.microsoft.com/Directory.Read.All'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when non-existent service principal specified in the APIs', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles') {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.microsoftGraphSp
          ]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return Promise.resolve({
          value: mocks.aadSp
        });
      }

      return Promise.reject(`Invalid GET request: ${opts.url}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                }
              ]
            }
          ],
          "spa": {
            "redirectUris": [
              "https://myspa.azurewebsites.net",
              "http://localhost:8080"
            ]
          },
          "web": {
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": true,
              "enableIdTokenIssuance": true
            }
          }
        })) {
        return Promise.resolve({
          "id": "f51ff52f-8f04-4924-91d0-636349eed65c",
          "deletedDateTime": null,
          "appId": "c505d465-9e4e-4bb4-b653-7b36d77cc94a",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:08:27.9188248Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": [
              "https://myspa.azurewebsites.net",
              "http://localhost:8080"
            ]
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
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                }
              ]
            }
          ],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": true,
              "enableIdTokenIssuance": true
            }
          }
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        platform: 'spa',
        apisDelegated: 'https://myapi.onmicrosoft.com/access_as_user',
        implicitFlow: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Service principal https://myapi.onmicrosoft.com not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when non-existent permission scope specified in the APIs', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles') {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.microsoftGraphSp
          ]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return Promise.resolve({
          value: mocks.aadSp
        });
      }

      return Promise.reject(`Invalid GET request: ${opts.url}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                }
              ]
            }
          ],
          "spa": {
            "redirectUris": [
              "https://myspa.azurewebsites.net",
              "http://localhost:8080"
            ]
          },
          "web": {
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": true,
              "enableIdTokenIssuance": true
            }
          }
        })) {
        return Promise.resolve({
          "id": "f51ff52f-8f04-4924-91d0-636349eed65c",
          "deletedDateTime": null,
          "appId": "c505d465-9e4e-4bb4-b653-7b36d77cc94a",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:08:27.9188248Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "spa": {
            "redirectUris": [
              "https://myspa.azurewebsites.net",
              "http://localhost:8080"
            ]
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
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
                  "type": "Scope"
                },
                {
                  "id": "06da0dbc-49e2-44d2-8312-53f166ab848a",
                  "type": "Scope"
                }
              ]
            }
          ],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": true,
              "enableIdTokenIssuance": true
            }
          }
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        platform: 'spa',
        apisDelegated: 'https://graph.microsoft.com/Read.Everything',
        implicitFlow: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Permission Read.Everything for service principal https://graph.microsoft.com not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when configuring secret failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issues GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "4d24b0c6-ad07-47c6-9bd8-9c167f9f758e",
          "deletedDateTime": null,
          "appId": "3c5bd51d-f1ac-4344-bd16-43396cadff14",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T14:58:18.7120335Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/4d24b0c6-ad07-47c6-9bd8-9c167f9f758e/addPassword') {
        return Promise.reject({
          error: {
            message: 'An error has occurred'
          }
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        withSecret: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when creating the AAD app reg failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issues GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(_ => Promise.reject({
      error: {
        message: 'An error has occurred'
      }
    }));

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when setting Application ID URI failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('Issued GET request'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject({
      error: {
        message: 'An error has occurred'
      }
    }));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg"
        })) {
        return Promise.resolve({
          "id": "c0e63919-057c-4e6b-be6c-8662e7aec4eb",
          "deletedDateTime": null,
          "appId": "b08d9318-5612-4f87-9f94-7414ef6f0c8a",
          "applicationTemplateId": null,
          "createdDateTime": "2020-12-31T19:14:23.9641082Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "M365x271534.onmicrosoft.com",
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
        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        uri: 'https://contoso.onmicrosoft.com/myapp'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates AAD app reg for a web app with service principal name with trailing slash', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles') {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27",
          "value": [
            mocks.mockCrmSp
          ]
        });
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames%2cappId%2coauth2PermissionScopes%2cappRoles&$skiptoken=X%274453707402000100000035536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D61323963386536336638613235536572766963655072696E636970616C5F34623131646566352D626561622D343232382D383835622D6132396338653633663861320000000000000000000000%27') {
        return Promise.resolve({
          value: mocks.aadSp
        });
      }

      return Promise.reject(`Invalid GET request: ${opts.url}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('Issued PATCH request'));
    sinon.stub(request, 'post').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My AAD app",
          "signInAudience": "AzureADMyOrg",
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000007-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "78ce3f0f-a1ce-49c2-8cde-64b5c0896db4",
                  "type": "Scope"
                }
              ]
            }
          ],
          "web": {
            "redirectUris": [
              "https://global.consent.azure-apim.net/redirect"
            ]
          }
        })) {
        return Promise.resolve({
          "id": "1cd23c5f-2cb4-4bd0-a582-d5b00f578dcd",
          "deletedDateTime": null,
          "appId": "702e65ba-cacb-4a2f-aa5c-e6460967bc20",
          "applicationTemplateId": null,
          "createdDateTime": "2021-02-21T09:44:05.953701Z",
          "displayName": "My AAD app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "optionalClaims": null,
          "publisherDomain": "m365404404.onmicrosoft.com",
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
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
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000007-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "78ce3f0f-a1ce-49c2-8cde-64b5c0896db4",
                  "type": "Scope"
                }
              ]
            }
          ],
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [
              "https://global.consent.azure-apim.net/redirect"
            ],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": []
          }

        });
      }

      return Promise.reject(`Invalid POST request: ${JSON.stringify(opts, null, 2)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        name: 'My AAD app',
        platform: 'web',
        redirectUris: 'https://global.consent.azure-apim.net/redirect',
        apisDelegated: 'https://admin.services.crm.dynamics.com/user_impersonation'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith({
          appId: '702e65ba-cacb-4a2f-aa5c-e6460967bc20',
          objectId: '1cd23c5f-2cb4-4bd0-a582-d5b00f578dcd',
          tenantId: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if specified platform value is not valid', () => {
    const actual = command.validate({ options: { name: 'My AAD app', platform: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if platform value is spa', () => {
    const actual = command.validate({ options: { name: 'My AAD app', platform: 'spa' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if platform value is web', () => {
    const actual = command.validate({ options: { name: 'My AAD app', platform: 'web' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if platform value is publicClient', () => {
    const actual = command.validate({ options: { name: 'My AAD app', platform: 'publicClient' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if redirectUris specified without platform', () => {
    const actual = command.validate({ options: { name: 'My AAD app', redirectUris: 'http://localhost:8080' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if redirectUris specified with platform', () => {
    const actual = command.validate({ options: { name: 'My AAD app', redirectUris: 'http://localhost:8080', platform: 'spa' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if scopeName specified without uri', () => {
    const actual = command.validate({ options: { name: 'My AAD app', scopeName: 'access_as_user', scopeAdminConsentDescription: 'Access as user', scopeAdminConsentDisplayName: 'Access as user' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scopeName specified without scopeAdminConsentDescription', () => {
    const actual = command.validate({ options: { name: 'My AAD app', scopeName: 'access_as_user', uri: 'https://contoso.onmicrosoft.com/myapp', scopeAdminConsentDisplayName: 'Access as user' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scopeName specified without scopeAdminConsentDisplayName', () => {
    const actual = command.validate({ options: { name: 'My AAD app', scopeName: 'access_as_user', uri: 'https://contoso.onmicrosoft.com/myapp', scopeAdminConsentDescription: 'Access as user' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if scopeName specified with uri, scopeAdminConsentDisplayName and scopeAdminConsentDescription', () => {
    const actual = command.validate({ options: { name: 'My AAD app', scopeName: 'access_as_user', uri: 'https://contoso.onmicrosoft.com/myapp', scopeAdminConsentDescription: 'Access as user', scopeAdminConsentDisplayName: 'Access as user' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified scopeConsentBy value is not valid', () => {
    const actual = command.validate({ options: { name: 'My AAD app', scopeConsentBy: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if scopeConsentBy is admins', () => {
    const actual = command.validate({ options: { name: 'My AAD app', scopeConsentBy: 'admins' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if scopeConsentBy is adminsAndUsers', () => {
    const actual = command.validate({ options: { name: 'My AAD app', scopeConsentBy: 'adminsAndUsers' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
