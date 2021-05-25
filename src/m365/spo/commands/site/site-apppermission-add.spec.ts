import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./site-apppermission-add');

describe(commands.SITE_APPPERMISSION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
      global.setTimeout
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
    assert.strictEqual(command.name.startsWith(commands.SITE_APPPERMISSION_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with an incorrect URL', (done) => {
    const actual = command.validate({
      options: {
        siteUrl: 'https;//contoso,sharepoint:com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = command.validate({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "123",
        appDisplayName: "Foo App"
      }
    });

    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and appDisplayName options are not specified', () => {
    const actual = command.validate({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation with a correct URL', (done) => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('fails validation if invalid value specified for permission', () => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "Invalid",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails when passing a site that does not exist', (done) => {
    const siteError = {
      "error": {
        "code": "itemNotFound",
        "message": "Requested site could not be found",
        "innerError": {
          "date": "2021-03-03T08:58:02",
          "request-id": "4e054f93-0eba-4743-be47-ce36b5f91120",
          "client-request-id": "dbd35b28-0ec3-6496-1279-0e1da3d028fe"
        }
      }
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('non-existing') === -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject(siteError);
    });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name-non-existing',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Requested site could not be found")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get Azure AD app when Azure AD app does not exists', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/myorganization/applications?$filter=') > -1) {
          return Promise.resolve({ value: [] });
        }
        return Promise.reject('The specified Azure AD app does not exist');
      });

    command.action(logger, {
      options: {
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified Azure AD app does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple Azure AD apps with same name exists', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/myorganization/applications') > -1) {
          return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications",
            "value": [
              {
                "id": "313f219e-b8a1-4454-84f0-ca05daa0fc4e",
                "deletedDateTime": null,
                "appId": "3166f9d8-f4e9-4b56-b634-dafcc9ecba8e",
                "applicationTemplateId": null,
                "createdDateTime": "2021-03-05T17:05:53Z",
                "displayName": "Foo App",
                "description": null,
                "groupMembershipClaims": null,
                "identifierUris": [],
                "isDeviceOnlyAuthSupported": null,
                "isFallbackPublicClient": null,
                "notes": null,
                "optionalClaims": null,
                "publisherDomain": "contoso.onmicrosoft.com",
                "signInAudience": "AzureADandPersonalMicrosoftAccount",
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
                  "requestedAccessTokenVersion": 2,
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
                "passwordCredentials": [
                  {
                    "customKeyIdentifier": null,
                    "displayName": "Foo App",
                    "endDateTime": "2299-12-31T00:00:00Z",
                    "hint": "Sl4",
                    "keyId": "85b90a55-0e86-4e2a-a1b5-889d6badb2ec",
                    "secretText": null,
                    "startDateTime": "2021-03-05T17:15:46.052Z"
                  },
                  {
                    "customKeyIdentifier": null,
                    "displayName": null,
                    "endDateTime": "2026-03-05T00:00:00Z",
                    "hint": "gwY",
                    "keyId": "0a67f4f2-67d5-446a-8b06-8fb84f699d16",
                    "secretText": null,
                    "startDateTime": "2021-03-05T17:05:55.9580541Z"
                  }
                ],
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
                },
                "spa": {
                  "redirectUris": []
                }
              },
              {
                "id": "8bb7fb05-64be-4b53-8936-f58d60946cf3",
                "deletedDateTime": null,
                "appId": "9bd7b7c0-e4a7-4b85-b0c6-20aaca0e25b7",
                "applicationTemplateId": null,
                "createdDateTime": "2021-03-24T14:43:35Z",
                "displayName": "Foo App",
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
                        "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
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
                    "enableAccessTokenIssuance": false,
                    "enableIdTokenIssuance": false
                  }
                },
                "spa": {
                  "redirectUris": []
                }
              }
            ]
          });
        }
        return Promise.reject('Multiple Azure AD app with displayName Foo App found: 3166f9d8-f4e9-4b56-b634-dafcc9ecba8e,9bd7b7c0-e4a7-4b85-b0c6-20aaca0e25b7');
      });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appDisplayName: "Foo App"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple Azure AD app with displayName Foo App found: 3166f9d8-f4e9-4b56-b634-dafcc9ecba8e,9bd7b7c0-e4a7-4b85-b0c6-20aaca0e25b7`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Adds an application permission to the site by appId', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/myorganization/applications') > -1) {
          return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications",
            "value": [
              {
                "id": "313f219e-b8a1-4454-84f0-ca05daa0fc4e",
                "deletedDateTime": null,
                "appId": "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
                "applicationTemplateId": null,
                "createdDateTime": "2021-03-05T17:05:53Z",
                "displayName": "Foo App",
                "description": null,
                "groupMembershipClaims": null,
                "identifierUris": [],
                "isDeviceOnlyAuthSupported": null,
                "isFallbackPublicClient": null,
                "notes": null,
                "optionalClaims": null,
                "publisherDomain": "contoso.onmicrosoft.com",
                "signInAudience": "AzureADandPersonalMicrosoftAccount",
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
                  "requestedAccessTokenVersion": 2,
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
                "passwordCredentials": [
                  {
                    "customKeyIdentifier": null,
                    "displayName": "Foo App",
                    "endDateTime": "2299-12-31T00:00:00Z",
                    "hint": "Sl4",
                    "keyId": "85b90a55-0e86-4e2a-a1b5-889d6badb2ec",
                    "secretText": null,
                    "startDateTime": "2021-03-05T17:15:46.052Z"
                  },
                  {
                    "customKeyIdentifier": null,
                    "displayName": null,
                    "endDateTime": "2026-03-05T00:00:00Z",
                    "hint": "gwY",
                    "keyId": "0a67f4f2-67d5-446a-8b06-8fb84f699d16",
                    "secretText": null,
                    "startDateTime": "2021-03-05T17:05:55.9580541Z"
                  }
                ],
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
                },
                "spa": {
                  "redirectUris": []
                }
              }]
          });
        }

        return Promise.reject('Invalid request');
      });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return Promise.resolve({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        output: "json"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Adds an application permission to the site by appDisplayName', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/myorganization/applications') > -1) {
          return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications",
            "value": [
              {
                "id": "313f219e-b8a1-4454-84f0-ca05daa0fc4e",
                "deletedDateTime": null,
                "appId": "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
                "applicationTemplateId": null,
                "createdDateTime": "2021-03-05T17:05:53Z",
                "displayName": "Foo App",
                "description": null,
                "groupMembershipClaims": null,
                "identifierUris": [],
                "isDeviceOnlyAuthSupported": null,
                "isFallbackPublicClient": null,
                "notes": null,
                "optionalClaims": null,
                "publisherDomain": "contoso.onmicrosoft.com",
                "signInAudience": "AzureADandPersonalMicrosoftAccount",
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
                  "requestedAccessTokenVersion": 2,
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
                "passwordCredentials": [
                  {
                    "customKeyIdentifier": null,
                    "displayName": "Foo App",
                    "endDateTime": "2299-12-31T00:00:00Z",
                    "hint": "Sl4",
                    "keyId": "85b90a55-0e86-4e2a-a1b5-889d6badb2ec",
                    "secretText": null,
                    "startDateTime": "2021-03-05T17:15:46.052Z"
                  },
                  {
                    "customKeyIdentifier": null,
                    "displayName": null,
                    "endDateTime": "2026-03-05T00:00:00Z",
                    "hint": "gwY",
                    "keyId": "0a67f4f2-67d5-446a-8b06-8fb84f699d16",
                    "secretText": null,
                    "startDateTime": "2021-03-05T17:05:55.9580541Z"
                  }
                ],
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
                },
                "spa": {
                  "redirectUris": []
                }
              }]
          });
        }

        return Promise.reject('Invalid request');
      });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return Promise.resolve({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appDisplayName: "Foo App",
        output: "json"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Adds an application permission to the site by appId and appDisplayName', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return Promise.resolve({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App",
        output: "json"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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
