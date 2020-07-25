import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./approleassignment-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as os from 'os';

describe(commands.APPROLEASSIGNMENT_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let getRequestStub = (): sinon.SinonStub => {

    return sinon.stub(request, 'get').callsFake((opts: any) => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6`) > -1) {

        // fake first call for getting service principal
        if (opts.url.indexOf('publisherName') === -1) {
          return Promise.resolve({ "value": [{ "objectType": "ServicePrincipal", "objectId": "57907bf8-73fa-43a6-89a5-1f603e29e452", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": ["isExplicit=False", "/subscriptions/fc3ca501-0a89-44cb-9192-5b60d3273989/resourcegroups/VM/providers/Microsoft.Compute/virtualMachines/VM2"], "appDisplayName": null, "appId": "1437b558-aa5c-48b2-9d6d-173e6dec518f", "applicationTemplateId": null, "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "VM2", "errorUrl": null, "homepage": null, "informationalUrls": null, "keyCredentials": [{ "customKeyIdentifier": "58967EE4A371AA7FE6D9EB74561B13184BEF0D95", "endDate": "2020-08-11T22:34:00Z", "keyId": "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "startDate": "2020-05-13T22:34:00Z", "type": "AsymmetricX509Cert", "usage": "Verify", "value": null }], "logoutUrl": null, "notificationEmailAddresses": [], "oauth2Permissions": [], "passwordCredentials": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyEndDateTime": null, "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "samlSingleSignOnSettings": null, "servicePrincipalNames": ["1437b558-aa5c-48b2-9d6d-173e6dec518f", "https://identity.azure.net/YWHbMwZ1ULadKOWfcXN7SEUXOloxXzsngLdjjJ8rGlg="], "servicePrincipalType": "ManagedIdentity", "signInAudience": null, "tags": [], "tokenEncryptionKeyId": null }] });
        }

        // second get request for searching for service principals by resource options value specified
        if (opts.url.indexOf('publisherName') !== -1) {
          return Promise.resolve({ "value": [{ "objectType": "ServicePrincipal", "objectId": "cd4f003c-d7cb-4245-9c59-a6997672a450", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": [], "appDisplayName": "Microsoft 365 SharePoint Online", "appId": "00000003-0000-0ff1-ce00-000000000000", "applicationTemplateId": null, "appOwnerTenantId": "f8cdef31-a31e-4b4a-93e4-5f571e91255a", "appRoleAssignmentRequired": false, "appRoles": [{ "allowedMemberTypes": ["Application"], "description": "Allows the app to create, read, update, and delete documents and list items in all site collections without a signed in user.", "displayName": "Read and write items in all site collections", "id": "fbcd29d2-fcca-4405-aded-518d457caae4", "isEnabled": true, "value": "Sites.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read documents and list items in all site collections without a signed in user.", "displayName": "Read items in all site collections", "id": "d13f72ca-a275-4b96-b789-48ebcc4da984", "isEnabled": true, "value": "Sites.Read.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to have full control of all site collections without a signed in user.", "displayName": "Have full control of all site collections", "id": "678536fe-1083-478a-9c59-b99265e6b0d3", "isEnabled": true, "value": "Sites.FullControl.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read, create, update, and delete document libraries and lists in all site collections without a signed in user.", "displayName": "Read and write items and lists in all site collections", "id": "9bff6588-13f2-4c48-bbf2-ddab62256b36", "isEnabled": true, "value": "Sites.Manage.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read enterprise managed metadata and to read basic site info without a signed in user.", "displayName": "Read managed metadata", "id": "2a8d57a5-4090-4a41-bf1c-3c621d2ccad3", "isEnabled": true, "value": "TermStore.Read.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to write enterprise managed metadata and to read basic site info without a signed in user.", "displayName": "Read and write managed metadata", "id": "c8e3537c-ec53-43b9-bed3-b2bd3617ae97", "isEnabled": true, "value": "TermStore.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read and update user profiles and to read basic site info without a signed in user.", "displayName": "Read and write user profiles", "id": "741f803b-c850-494e-b5df-cde7c675a1ca", "isEnabled": true, "value": "User.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read user profiles without a signed in user.", "displayName": "Read user profiles", "id": "df021288-bdef-4463-88db-98f22de89214", "isEnabled": true, "value": "User.Read.All" }], "displayName": "Microsoft 365 SharePoint Online", "errorUrl": null, "homepage": null, "informationalUrls": { "termsOfService": null, "support": null, "privacy": null, "marketing": null }, "keyCredentials": [], "logoutUrl": "https://signout.sharepoint.com/_layouts/15/expirecookies.aspx", "notificationEmailAddresses": [], "oauth2Permissions": [{ "adminConsentDescription": "Allows the app to read managed metadata and to read basic site info on behalf of the signed-in user.", "adminConsentDisplayName": "Read managed metadata", "id": "a468ea40-458c-4cc2-80c4-51781af71e41", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to read managed metadata and to read basic site info on your behalf.", "userConsentDisplayName": "Read managed metadata", "value": "TermStore.Read.All" }, { "adminConsentDescription": "Allows the app to read, create, update, and delete managed metadata and to read basic site info on behalf of the signed-in user.", "adminConsentDisplayName": "Read and write managed metadata", "id": "59a198b5-0420-45a8-ae59-6da1cb640505", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to read, create, update, and delete managed metadata and to read basic site info on your behalf.", "userConsentDisplayName": "Read and write managed metadata", "value": "TermStore.ReadWrite.All" }, { "adminConsentDescription": "Allows the app to run search queries and to read basic site info on behalf of the current signed-in user. Search results are based on the user's permissions instead of the app's permissions.", "adminConsentDisplayName": "Run search queries as a user", "id": "1002502a-9a71-4426-8551-69ab83452fab", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to run search queries and to read basic site info on your behalf. Search results are based on your permissions.", "userConsentDisplayName": "Run search queries ", "value": "Sites.Search.All" }, { "adminConsentDescription": "Allows the app to read documents and list items in all site collections on behalf of the signed-in user.", "adminConsentDisplayName": "Read items in all site collections", "id": "4e0d77b0-96ba-4398-af14-3baa780278f4", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read documents and list items in all site collections on your behalf.", "userConsentDisplayName": "Read items in all site collections", "value": "AllSites.Read" }, { "adminConsentDescription": "Allows the app to create, read, update, and delete documents and list items in all site collections on behalf of the signed-in user.", "adminConsentDisplayName": "Read and write items in all site collections", "id": "640ddd16-e5b7-4d71-9690-3f4022699ee7", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to create, read, update, and delete documents and list items in all site collections on your behalf.", "userConsentDisplayName": "Read and write items in all site collections", "value": "AllSites.Write" }, { "adminConsentDescription": "Allows the app to read, create, update, and delete document libraries and lists in all site collections on behalf of the signed-in user.", "adminConsentDisplayName": "Read and write items and lists in all site collections", "id": "b3f70a70-8a4b-4f95-9573-d71c496a53f4", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read, create, update, and delete document libraries and lists in all site collections on your behalf.", "userConsentDisplayName": "Read and write items and lists in all site collections", "value": "AllSites.Manage" }, { "adminConsentDescription": "Allows the app to have full control of all site collections on behalf of the signed-in user.", "adminConsentDisplayName": "Have full control of all site collections", "id": "56680e0d-d2a3-4ae1-80d8-3c4f2100e3d0", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to have full control of all site collections on your behalf.", "userConsentDisplayName": "Have full control of all site collections", "value": "AllSites.FullControl" }, { "adminConsentDescription": "Allows the app to read the current user's files.", "adminConsentDisplayName": "Read user files", "id": "dd2c8d78-58e1-46d7-82dd-34d411282686", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read your files.", "userConsentDisplayName": "Read your files", "value": "MyFiles.Read" }, { "adminConsentDescription": "Allows the app to read, create, update, and delete the current user's files.", "adminConsentDisplayName": "Read and write user files", "id": "2cfdc887-d7b4-4798-9b33-3d98d6b95dd2", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read, create, update, and delete your files.", "userConsentDisplayName": "Read and write your files", "value": "MyFiles.Write" }, { "adminConsentDescription": "Allows the app to read and update user profiles and to read basic site info on behalf of the signed-in user.", "adminConsentDisplayName": "Read and write user profiles", "id": "82866913-39a9-4be7-8091-f4fa781088ae", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to read and update user profiles and to read basic site info on your behalf.", "userConsentDisplayName": "Read and write user profiles", "value": "User.ReadWrite.All" }, { "adminConsentDescription": "Allows the app to read user profiles and to read basic site info on behalf of the signed-in user.", "adminConsentDisplayName": "Read user profiles", "id": "0cea5a30-f6f8-42b5-87a0-84cc26822e02", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to read user profiles and basic site info on your behalf.", "userConsentDisplayName": "Read user profiles", "value": "User.Read.All" }], "passwordCredentials": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyEndDateTime": null, "preferredTokenSigningKeyThumbprint": null, "publisherName": "Microsoft Services", "replyUrls": [], "samlMetadataUrl": null, "samlSingleSignOnSettings": null, "servicePrincipalNames": ["00000003-0000-0ff1-ce00-000000000000/*.sharepoint.com", "00000003-0000-0ff1-ce00-000000000000", "https://microsoft.sharepoint-df.com"], "servicePrincipalType": "Application", "signInAudience": "AzureADMultipleOrgs", "tags": [], "tokenEncryptionKeyId": null }] });
        }
      }
      return Promise.reject();
    })
  }

  let postRequestStub = (): sinon.SinonStub => {

    return sinon.stub(request, 'post').callsFake((opts: any) => {
      return Promise.resolve({ "objectType": "AppRoleAssignment", "objectId": "-HuQV_pzpkOJpR9gPinkUtHqOUvb9FRKurxnugdMSPs", "deletionTimestamp": null, "creationTimestamp": "2020-05-15T09:54:37.2055435Z", "id": "fff194f1-7dce-4428-8301-1badb5518201", "principalDisplayName": "VM2", "principalId": "57907bf8-73fa-43a6-89a5-1f603e29e452", "principalType": "ServicePrincipal", "resourceDisplayName": "Microsoft Graph", "resourceId": "1a3413b4-c588-45db-a77f-da44a564c495" });
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get,
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
    assert.strictEqual(command.name.startsWith(commands.APPROLEASSIGNMENT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets App Role assignments for service principal with specified displayName', (done) => {
    getRequestStub();
    postRequestStub();

    cmdInstance.action({ options: { displayName: 'myapp', resource: 'SharePoint', scope: 'Sites.Read.All' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].objectId, '-HuQV_pzpkOJpR9gPinkUtHqOUvb9FRKurxnugdMSPs');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].principalDisplayName, 'VM2');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].resourceDisplayName, 'Microsoft Graph');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets App Role assignments for service principal with specified objectId and multiple scopes', (done) => {
    getRequestStub();
    postRequestStub();

    cmdInstance.action({ options: { objectId: '77907bf8-73fa-43a6-89a5-1f603e29e452', resource: 'SharePoint', scope: 'Sites.Read.All,Sites.ReadWrite.All' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].objectId, '-HuQV_pzpkOJpR9gPinkUtHqOUvb9FRKurxnugdMSPs');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].principalDisplayName, 'VM2');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].resourceDisplayName, 'Microsoft Graph');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][1].objectId, '-HuQV_pzpkOJpR9gPinkUtHqOUvb9FRKurxnugdMSPs');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][1].principalDisplayName, 'VM2');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][1].resourceDisplayName, 'Microsoft Graph');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets App Role assignments for service principal with specified displayName and output json', (done) => {
    getRequestStub();
    postRequestStub();

    cmdInstance.action({ options: { displayName: 'myapp', resource: 'SharePoint', scope: 'Sites.Read.All', output: 'json' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].objectId, '-HuQV_pzpkOJpR9gPinkUtHqOUvb9FRKurxnugdMSPs');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].principalDisplayName, 'VM2');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].resourceDisplayName, 'Microsoft Graph');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 'fff194f1-7dce-4428-8301-1badb5518201');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].principalId, '57907bf8-73fa-43a6-89a5-1f603e29e452');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].principalType, 'ServicePrincipal');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].resourceId, '1a3413b4-c588-45db-a77f-da44a564c495');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets App Role assignments for service principal with specified appId (debug)', (done) => {
    getRequestStub();
    postRequestStub();

    cmdInstance.action({ options: { debug: true, appId: 'fff194f1-7dce-4428-8301-1badb5518201', resource: 'SharePoint', scope: 'Sites.Read.All' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE') !== 1, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles intune alias for the resource option value', (done) => {
    getRequestStub();
    postRequestStub();

    cmdInstance.action({ options: { debug: true, appId: 'fff194f1-7dce-4428-8301-1badb5518201', resource: 'intune', scope: 'Sites.Read.All' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE') !== 1, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles exchange alias for the resource option value', (done) => {
    getRequestStub();
    postRequestStub();

    cmdInstance.action({ options: { debug: true, appId: 'fff194f1-7dce-4428-8301-1badb5518201', resource: 'exchange', scope: 'Sites.Read.All' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE') !== 1, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles appId for the resource option value', (done) => {
    getRequestStub();
    postRequestStub();

    cmdInstance.action({ options: { debug: true, appId: 'fff194f1-7dce-4428-8301-1badb5518201', resource: 'fff194f1-7dce-4428-8301-1badb5518201', scope: 'Sites.Read.All' } }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE') !== 1, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('rejects if not app roles found for the specified resource option value', (done) => {
    postRequestStub();
    sinon.stub(request, 'get').callsFake((opts: any): Promise<any> => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6`) > -1) {

        // fake first call for getting service principal
        if (opts.url.indexOf('publisherName') === -1) {
          return Promise.resolve({ "value": [{ "objectType": "ServicePrincipal", "objectId": "57907bf8-73fa-43a6-89a5-1f603e29e452", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": ["isExplicit=False", "/subscriptions/fc3ca501-0a89-44cb-9192-5b60d3273989/resourcegroups/VM/providers/Microsoft.Compute/virtualMachines/VM2"], "appDisplayName": null, "appId": "1437b558-aa5c-48b2-9d6d-173e6dec518f", "applicationTemplateId": null, "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "VM2", "errorUrl": null, "homepage": null, "informationalUrls": null, "keyCredentials": [{ "customKeyIdentifier": "58967EE4A371AA7FE6D9EB74561B13184BEF0D95", "endDate": "2020-08-11T22:34:00Z", "keyId": "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "startDate": "2020-05-13T22:34:00Z", "type": "AsymmetricX509Cert", "usage": "Verify", "value": null }], "logoutUrl": null, "notificationEmailAddresses": [], "oauth2Permissions": [], "passwordCredentials": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyEndDateTime": null, "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "samlSingleSignOnSettings": null, "servicePrincipalNames": ["1437b558-aa5c-48b2-9d6d-173e6dec518f", "https://identity.azure.net/YWHbMwZ1ULadKOWfcXN7SEUXOloxXzsngLdjjJ8rGlg="], "servicePrincipalType": "ManagedIdentity", "signInAudience": null, "tags": [], "tokenEncryptionKeyId": null }] });
        }

        // second get request for searching for service principals by resource options value specified
        if (opts.url.indexOf('publisherName') !== -1) {
          return Promise.resolve({ "value": [{ objectId: "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "appRoles": [] }] });
        }
      }
      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: true, appId: 'fff194f1-7dce-4428-8301-1badb5518201', resource: 'SharePoint', scope: 'Sites.Read.All' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The resource 'SharePoint' does not have any application permissions available.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('rejects if not app roles found for the specified resource option value', (done) => {
    postRequestStub();
    sinon.stub(request, 'get').callsFake((opts: any): Promise<any> => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6`) > -1) {

        // fake first call for getting service principal
        if (opts.url.indexOf('publisherName') === -1) {
          return Promise.resolve({ "value": [{ "objectType": "ServicePrincipal", "objectId": "57907bf8-73fa-43a6-89a5-1f603e29e452", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": ["isExplicit=False", "/subscriptions/fc3ca501-0a89-44cb-9192-5b60d3273989/resourcegroups/VM/providers/Microsoft.Compute/virtualMachines/VM2"], "appDisplayName": null, "appId": "1437b558-aa5c-48b2-9d6d-173e6dec518f", "applicationTemplateId": null, "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "VM2", "errorUrl": null, "homepage": null, "informationalUrls": null, "keyCredentials": [{ "customKeyIdentifier": "58967EE4A371AA7FE6D9EB74561B13184BEF0D95", "endDate": "2020-08-11T22:34:00Z", "keyId": "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "startDate": "2020-05-13T22:34:00Z", "type": "AsymmetricX509Cert", "usage": "Verify", "value": null }], "logoutUrl": null, "notificationEmailAddresses": [], "oauth2Permissions": [], "passwordCredentials": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyEndDateTime": null, "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "samlSingleSignOnSettings": null, "servicePrincipalNames": ["1437b558-aa5c-48b2-9d6d-173e6dec518f", "https://identity.azure.net/YWHbMwZ1ULadKOWfcXN7SEUXOloxXzsngLdjjJ8rGlg="], "servicePrincipalType": "ManagedIdentity", "signInAudience": null, "tags": [], "tokenEncryptionKeyId": null }] });
        }

        // second get request for searching for service principals by resource options value specified
        if (opts.url.indexOf('publisherName') !== -1) {
          return Promise.resolve({ "value": [{ objectId: "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "appRoles": [{ value: 'Scope1', id: '1' }, { value: 'Scope2', id: '2' }] }] });
        }
      }
      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: true, appId: 'fff194f1-7dce-4428-8301-1badb5518201', resource: 'SharePoint', scope: 'Sites.Read.All' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The scope value 'Sites.Read.All' you have specified does not exist for SharePoint. ${os.EOL}Available scopes (application permissions) are: ${os.EOL}Scope1${os.EOL}Scope2`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('rejects if more than one service principal found', (done) => {
    postRequestStub();
    sinon.stub(request, 'get').callsFake((opts: any): Promise<any> => {

      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6`) > -1) {

        // fake first call for getting service principal
        if (opts.url.indexOf('publisherName') === -1) {
          return Promise.resolve({ "value": [{ "objectType": "ServicePrincipal", "objectId": "57907bf8-73fa-43a6-89a5-1f603e29e452", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": ["isExplicit=False", "/subscriptions/fc3ca501-0a89-44cb-9192-5b60d3273989/resourcegroups/VM/providers/Microsoft.Compute/virtualMachines/VM2"], "appDisplayName": null, "appId": "1437b558-aa5c-48b2-9d6d-173e6dec518f", "applicationTemplateId": null, "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "VM2", "errorUrl": null, "homepage": null, "informationalUrls": null, "keyCredentials": [{ "customKeyIdentifier": "58967EE4A371AA7FE6D9EB74561B13184BEF0D95", "endDate": "2020-08-11T22:34:00Z", "keyId": "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "startDate": "2020-05-13T22:34:00Z", "type": "AsymmetricX509Cert", "usage": "Verify", "value": null }], "logoutUrl": null, "notificationEmailAddresses": [], "oauth2Permissions": [], "passwordCredentials": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyEndDateTime": null, "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "samlSingleSignOnSettings": null, "servicePrincipalNames": ["1437b558-aa5c-48b2-9d6d-173e6dec518f", "https://identity.azure.net/YWHbMwZ1ULadKOWfcXN7SEUXOloxXzsngLdjjJ8rGlg="], "servicePrincipalType": "ManagedIdentity", "signInAudience": null, "tags": [], "tokenEncryptionKeyId": null }, { "objectType": "ServicePrincipal", "objectId": "57907bf8-73fa-43a6-89a5-1f603e29e452", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": ["isExplicit=False", "/subscriptions/fc3ca501-0a89-44cb-9192-5b60d3273989/resourcegroups/VM/providers/Microsoft.Compute/virtualMachines/VM2"], "appDisplayName": null, "appId": "1437b558-aa5c-48b2-9d6d-173e6dec518f", "applicationTemplateId": null, "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "VM2", "errorUrl": null, "homepage": null, "informationalUrls": null, "keyCredentials": [{ "customKeyIdentifier": "58967EE4A371AA7FE6D9EB74561B13184BEF0D95", "endDate": "2020-08-11T22:34:00Z", "keyId": "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "startDate": "2020-05-13T22:34:00Z", "type": "AsymmetricX509Cert", "usage": "Verify", "value": null }], "logoutUrl": null, "notificationEmailAddresses": [], "oauth2Permissions": [], "passwordCredentials": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyEndDateTime": null, "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "samlSingleSignOnSettings": null, "servicePrincipalNames": ["1437b558-aa5c-48b2-9d6d-173e6dec518f", "https://identity.azure.net/YWHbMwZ1ULadKOWfcXN7SEUXOloxXzsngLdjjJ8rGlg="], "servicePrincipalType": "ManagedIdentity", "signInAudience": null, "tags": [], "tokenEncryptionKeyId": null }] });
        }

        // second get request for searching for service principals by resource options value specified
        if (opts.url.indexOf('publisherName') !== -1) {
          return Promise.resolve({ "value": [{ objectId: "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "appRoles": [{ value: 'Scope1', id: '1' }, { value: 'Scope2', id: '2' }] }] });
        }
      }
      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: true, appId: 'fff194f1-7dce-4428-8301-1badb5518201', resource: 'SharePoint', scope: 'Sites.Read.All' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("More than one service principal found. Please use the appId or objectId option to make sure the right service principal is specified.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, appId: '36e3a540-6f25-4483-9542-9f5fa00bb633' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither appId, objectId nor displayName are not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '123', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { objectId: '123', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '123', displayName: 'abc', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both objectId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { objectId: '123', displayName: 'abc', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both objectId, appId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '123', objectId: '123', displayName: 'abc', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  })

  it('passes validation when the appId option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '57907bf8-73fa-43a6-89a5-1f603e29e452', resource: 'abc', scope: 'abc' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying appId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying displayName', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});

