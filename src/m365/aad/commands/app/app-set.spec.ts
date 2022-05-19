import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./app-set');

describe(commands.APP_SET, () => {

  //#region Mocked Responses  
  const appDetailsResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity", "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83", "deletedDateTime": null, "appId": "ff254847-12c7-44cf-921e-8883dbd622a7", "applicationTemplateId": null, "disabledByMicrosoftStatus": null, "createdDateTime": "2022-02-07T08:51:18Z", "displayName": "Angular Teams app", "description": null, "groupMembershipClaims": null, "identifierUris": ["api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"], "isDeviceOnlyAuthSupported": null, "isFallbackPublicClient": null, "notes": null, "publisherDomain": "contoso.onmicrosoft.com", "serviceManagementReference": null, "signInAudience": "AzureADMyOrg", "tags": [], "tokenEncryptionKeyId": null, "defaultRedirectUri": null, "certification": null, "optionalClaims": null, "addIns": [], "api": { "acceptMappedClaims": null, "knownClientApplications": [], "requestedAccessTokenVersion": null, "oauth2PermissionScopes": [{ "adminConsentDescription": "Access as a user", "adminConsentDisplayName": "Access as a user", "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5", "isEnabled": true, "type": "User", "userConsentDescription": null, "userConsentDisplayName": null, "value": "access_as_user" }], "preAuthorizedApplications": [{ "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346", "delegatedPermissionIds": ["cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"] }, { "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264", "delegatedPermissionIds": ["cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"] }] }, "appRoles": [], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "parentalControlSettings": { "countriesBlockedForMinors": [], "legalAgeGroupRule": "Allow" }, "passwordCredentials": [], "publicClient": { "redirectUris": [] }, "requiredResourceAccess": [{ "resourceAppId": "00000003-0000-0000-c000-000000000000", "resourceAccess": [{ "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope" }] }], "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "web": { "homePageUrl": null, "logoutUrl": null, "redirectUris": [], "implicitGrantSettings": { "enableAccessTokenIssuance": false, "enableIdTokenIssuance": false } }, "spa": { "redirectUris": [] } };
  //#endregion

  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates uri for the specified appId', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq 'bc724b77-da87-43a9-b385-6ebaaf969db8'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates uri for the specified objectId', (done) => {
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates uri for the specified name', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'My app',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips updating uri if no uri specified', (done) => {
    command.action(logger, {
      options: {
        debug: false,
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets spa redirectUri for an app without redirectUris', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/e4528262-097a-42eb-98e1-19f073dbee45`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "e4528262-097a-42eb-98e1-19f073dbee45",
          "deletedDateTime": null,
          "appId": "842e1a6f-7492-4b7d-8278-563036f5bd39",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-10T08:01:06Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5196.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
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
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
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

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/e4528262-097a-42eb-98e1-19f073dbee45' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": []
          },
          "spa": {
            "redirectUris": [
              "https://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth"
            ]
          },
          "web": {
            "redirectUris": []
          }
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        objectId: 'e4528262-097a-42eb-98e1-19f073dbee45',
        redirectUris: 'https://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth',
        platform: 'spa'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets web redirectUri for an app with existing spa redirectUris', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
          "deletedDateTime": null,
          "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-07T08:51:18Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
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
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          }
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": []
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          },
          "web": {
            "redirectUris": [
              "https://foo.com"
            ]
          }
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        redirectUris: 'https://foo.com',
        platform: 'web'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets publicClient redirectUri for an app with existing spa and web redirectUris', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
          "deletedDateTime": null,
          "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-07T08:51:18Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
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
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [
              "https://foo.com"
            ],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          }
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": [
              "https://foo1.com"
            ]
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          },
          "web": {
            "redirectUris": [
              "https://foo.com"
            ]
          }
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        redirectUris: 'https://foo1.com',
        platform: 'publicClient'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('replaces existing redirectUri with a new one', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
          "deletedDateTime": null,
          "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-07T08:51:18Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
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
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          }
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": []
          },
          "spa": {
            "redirectUris": [
              "http://localhost/auth",
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth"
            ]
          },
          "web": {
            "redirectUris": []
          }
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        redirectUris: 'https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth',
        platform: 'spa',
        redirectUrisToRemove: 'https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new certificate using base64 string', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return Promise.resolve(appDetailsResponse);
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "keyCredentials": []
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "keyCredentials": [{
            "type": "AsymmetricX509Cert",
            "usage": "Verify",
            "displayName": "some certificate",
            "key": "somecertificatebase64string"
          }]
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateBase64Encoded: 'somecertificatebase64string'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new certificate using base64 string (with null keyCredentials)', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return Promise.resolve(appDetailsResponse);
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "keyCredentials": null
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "keyCredentials": [{
            "type": "AsymmetricX509Cert",
            "usage": "Verify",
            "displayName": "some certificate",
            "key": "somecertificatebase64string"
          }]
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateBase64Encoded: 'somecertificatebase64string'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new certificate using certificate file', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return Promise.resolve(appDetailsResponse);
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "keyCredentials": []
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "keyCredentials": [{
            "type": "AsymmetricX509Cert",
            "usage": "Verify",
            "displayName": "some certificate",
            "key": "somecertificatebase64string"
          }]
        })) {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => "somecertificatebase64string");

    command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateFile: 'C:\\temp\\some-certificate.cer'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when certificate file cannot be read', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return Promise.resolve(appDetailsResponse);
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => { throw new Error("An error has occurred"); });


    command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateFile: 'C:\\temp\\some-certificate.cer'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Error reading certificate file: Error: An error has occurred. Please add the certificate using base64 option '--certificateBase64Encoded'.`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with objectId not found', (done) => {
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));

    command.action(logger, {
      options: {
        debug: false,
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with the appId not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with name not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        name: 'My app',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No Azure AD application registration with name My app found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when multiple apps with the specified name found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        name: 'My app',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple Azure AD application registration with name My app found. Please disambiguate (app object IDs): 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving information about app through appId failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `An error has occurred`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving information about app through name failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        name: 'My app',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `An error has occurred`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if appId and objectId specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and name specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and name specified', () => {
    const actual = command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, objectId nor name specified', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if redirectUris specified without platform', () => {
    const actual = command.validate({ options: { objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', redirectUris: 'https://foo.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid platform specified', () => {
    const actual = command.validate({ options: { objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', redirectUris: 'https://foo.com', platform: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if certificateDisplayName is specified without certificate', () => {
    const actual = command.validate({ options: { objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', certificateDisplayName: 'Some certificate' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both certificateBase64Encoded and certificateFile are specified', () => {
    const actual = command.validate({ options: { objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', certificateFile: 'c:\\temp\\some-certificate.cer', certificateBase64Encoded: 'somebase64string' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if certificateFile specified with certificateDisplayName', () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);

    const actual = command.validate({ options: { name: 'My AAD app', certificateDisplayName: 'Some certificate', certificateFile: 'c:\\temp\\some-certificate.cer' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when certificate file is not found', () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);

    const actual = command.validate({ options: { debug: true, objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83', certificateDisplayName: 'some certificate', certificateFile: 'C:\\temp\\some-certificate.cer' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (appId)', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (objectId)', () => {
    const actual = command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', () => {
    const actual = command.validate({ options: { name: 'My app', uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when redirectUris specified with spa', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'spa' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when redirectUris specified with publicClient', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'publicClient' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when redirectUris specified with web', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'web' } });
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
