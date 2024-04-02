import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import { entraApp } from './entraApp.js';
import { sinonUtil } from "./sinonUtil.js";
import { cli } from '../cli/cli.js';

describe('utils/entraApp', () => {
  beforeEach(() => {
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      request.post,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('should create a new application successfully', async () => {
    const mockApplicationInfo = { displayName: "My Microsoft Entra app" };
    const expectedResponse = {
      "id": "9b1e2c08-6e35-4134-a0ac-16ab154cd05a",
      "deletedDateTime": null,
      "appId": "62f0f128-987f-47f2-827a-be50d0d894c7",
      "applicationTemplateId": null,
      "createdDateTime": "2020-12-31T14:50:40.1806422Z",
      "displayName": "My Microsoft Entra app",
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
    };

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app"
        })) {
        return expectedResponse;
      }

      return `Invalid Request ${JSON.stringify(opts, null, 2)}`;
    });

    const result = await entraApp.createEntraApp(mockApplicationInfo);
    assert.deepStrictEqual(result, expectedResponse);
  });

  it('should handle errors when creating an application', async () => {
    const mockApplicationInfo = { /* Mock application info */ };
    const expectedError = new Error('Failed to create application');

    sinon.stub(request, 'post').rejects(expectedError);

    await assert.rejects(async () => {
      await entraApp.createEntraApp(mockApplicationInfo);
    }, expectedError);
  });

  it('should update an existing application successfully', async () => {
    const appId = 'mockAppId';
    const mockApplicationInfo = { displayName: "My Microsoft Entra app" };
    const expectedResponse = {
      "id": "9b1e2c08-6e35-4134-a0ac-16ab154cd05a",
      "deletedDateTime": null,
      "appId": "62f0f128-987f-47f2-827a-be50d0d894c7",
      "applicationTemplateId": null,
      "createdDateTime": "2020-12-31T14:50:40.1806422Z",
      "displayName": "My Microsoft Entra app",
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
    };

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${appId}` &&
        JSON.stringify(opts.data) === JSON.stringify({
          "displayName": "My Microsoft Entra app"
        })) {
        return expectedResponse;
      }

      return `Invalid Request ${JSON.stringify(opts, null, 2)}`;
    });

    const result = await entraApp.updateEntraApp(appId, mockApplicationInfo);
    assert.deepStrictEqual(result, expectedResponse);
  });

  it('should handle errors when updating an application', async () => {
    const appId = 'mockAppId';
    const mockApplicationInfo = { /* Mock application info */ };
    const expectedError = new Error('Failed to update application');

    sinon.stub(request, 'patch').rejects(expectedError);

    await assert.rejects(async () => {
      await entraApp.updateEntraApp(appId, mockApplicationInfo);
    }, expectedError);
  });

  it('should add a role to a service principal successfully', async () => {
    const objectId = 'mockObjectId';
    const resourceId = 'mockResourceId';
    const appRoleId = 'mockAppRoleId';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/${objectId}/appRoleAssignments` &&
        JSON.stringify(opts.data) === JSON.stringify({
          data: {
            appRoleId: appRoleId,
            principalId: objectId,
            resourceId: resourceId
          }
        })) {
        return;
      }

      return `Invalid Request ${JSON.stringify(opts, null, 2)}`;
    });

    await entraApp.addRoleToServicePrincipal(objectId, resourceId, appRoleId);
  });

  it('should handle errors when adding a role to a service principal', async () => {
    // Mock data and expected error
    const objectId = 'mockObjectId';
    const resourceId = 'mockResourceId';
    const appRoleId = 'mockAppRoleId';
    const expectedError = new Error('Failed to add role to service principal');

    sinon.stub(request, 'post').rejects(expectedError);

    await assert.rejects(async () => {
      await entraApp.addRoleToServicePrincipal(objectId, resourceId, appRoleId);
    }, expectedError);
  });

  it('should grant OAuth2 permission successfully', async () => {
    const appId = 'mockAppId';
    const resourceId = 'mockResourceId';
    const scopeName = 'mockScopeName';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants` &&
        JSON.stringify(opts.data) === JSON.stringify({
          data: {
            clientId: appId,
            consentType: "AllPrincipals",
            principalId: null,
            resourceId: resourceId,
            scope: scopeName
          }
        })) {
        return;
      }

      return `Invalid Request ${JSON.stringify(opts, null, 2)}`;
    });

    await entraApp.grantOAuth2Permission(appId, resourceId, scopeName);
  });

  it('should handle errors when granting OAuth2 permission', async () => {
    const appId = 'mockAppId';
    const resourceId = 'mockResourceId';
    const scopeName = 'mockScopeName';
    const expectedError = new Error('Failed to grant OAuth2 permission');

    sinon.stub(request, 'post').rejects(expectedError);

    await assert.rejects(async () => {
      await entraApp.grantOAuth2Permission(appId, resourceId, scopeName);
    }, expectedError);
  });

  it('should create a new service principal successfully', async () => {
    const appId = 'mockAppId';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/servicePrincipals` &&
        JSON.stringify(opts.data) === JSON.stringify({
          data: {
            appId: appId
          }
        })) {
        return;
      }

      return `Invalid Request ${JSON.stringify(opts, null, 2)}`;
    });

    await entraApp.createServicePrincipal(appId);
  });

  it('should handle errors when creating a service principal', async () => {
    const appId = 'mockAppId';
    const expectedError = new Error('Failed to create service principal');

    sinon.stub(request, 'post').rejects(expectedError);

    await assert.rejects(async () => {
      await entraApp.createServicePrincipal(appId);
    }, expectedError);
  });
});