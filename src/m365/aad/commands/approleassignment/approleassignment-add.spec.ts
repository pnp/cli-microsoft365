import * as assert from 'assert';
import * as os from 'os';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./approleassignment-add');

describe(commands.APPROLEASSIGNMENT_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const getRequestStub = (): sinon.SinonStub => {

    return sinon.stub(request, 'get').callsFake((opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?`) > -1) {
        // fake first call for getting service principal
        if (opts.url.indexOf('startswith') === -1) {
          return Promise.resolve({ "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals", "value": [{ "id": "24448e9c-d0fa-43d1-a1dd-e279720969a0", "deletedDateTime": null, "accountEnabled": true, "alternativeNames": [], "appDisplayName": "myapp", "appDescription": null, "appId": "26e49d05-4227-4ace-ae52-9b8f08f37184", "applicationTemplateId": null, "appOwnerOrganizationId": "c8e571e1-d528-43d9-8776-dc51157d615a", "appRoleAssignmentRequired": false, "createdDateTime": "2020-08-29T18:35:13Z", "description": null, "displayName": "myapp", "homepage": null, "loginUrl": null, "logoutUrl": null, "notes": null, "notificationEmailAddresses": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyThumbprint": null, "replyUrls": ["https://login.microsoftonline.com/common/oauth2/nativeclient"], "resourceSpecificApplicationPermissions": [], "samlSingleSignOnSettings": null, "servicePrincipalNames": ["26e49d05-4227-4ace-ae52-9b8f08f37184"], "servicePrincipalType": "Application", "signInAudience": "AzureADMyOrg", "tags": ["WindowsAzureActiveDirectoryIntegratedApp"], "tokenEncryptionKeyId": null, "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "addIns": [], "appRoles": [], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "oauth2PermissionScopes": [], "passwordCredentials": [] }] });
        }
        // second get request for searching for service principals by resource options value specified
        return Promise.resolve({ "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals", "value": [{ "id": "df3d00f0-a24d-45a9-ba8b-3b0934ec3a6c", "deletedDateTime": null, "accountEnabled": true, "alternativeNames": [], "appDisplayName": "Office 365 SharePoint Online", "appDescription": null, "appId": "00000003-0000-0ff1-ce00-000000000000", "applicationTemplateId": null, "appOwnerOrganizationId": "f8cdef31-a31e-4b4a-93e4-5f571e91255a", "appRoleAssignmentRequired": false, "createdDateTime": "2019-01-11T07:34:21Z", "description": null, "displayName": "Office 365 SharePoint Online", "homepage": null, "loginUrl": null, "logoutUrl": "https://signout.sharepoint.com/_layouts/15/expirecookies.aspx", "notes": null, "notificationEmailAddresses": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyThumbprint": null, "replyUrls": ["https://www180015.066dapp.com/_layouts/15/spolanding.aspx", "https://www167017.080dapp.com/_layouts/15/spolanding.aspx", "https://www174015.019dapp.com/_layouts/15/spolanding.aspx", "https://www162015.079dapp.com/_layouts/15/spolanding.aspx", "https://www156015.077dapp.com/_layouts/15/spolanding.aspx", "https://www158015.075dapp.com/_layouts/15/spolanding.aspx", "https://www145007.074dapp.com/_layouts/15/spolanding.aspx", "https://www148015.030dapp.com/_layouts/15/spolanding.aspx", "https://www141017.028dapp.com/_layouts/15/spolanding.aspx", "https://www143020.025dapp.com/_layouts/15/spolanding.aspx", "https://www138015.076dapp.com/_layouts/15/spolanding.aspx", "https://www136028.062dapp.com/_layouts/15/spolanding.aspx", "https://www129017.072dapp.com/_layouts/15/spolanding.aspx", "https://www127017.005dapp.com/_layouts/15/spolanding.aspx", "https://www124016.032dapp.com/_layouts/15/spolanding.aspx", "https://www117017.063dapp.com/_layouts/15/spolanding.aspx", "https://www115014.071dapp.com/_layouts/15/spolanding.aspx", "https://www111031.045dapp.com/_layouts/15/spolanding.aspx", "https://www113025.044dapp.com/_layouts/15/spolanding.aspx", "https://www105021.059dapp.com/_layouts/15/spolanding.aspx", "https://www92050.065dapp.com/_layouts/15/spolanding.aspx", "https://www32058.050dapp.com/_layouts/15/spolanding.aspx", "https://www29079.048dapp.com/_layouts/15/spolanding.aspx", "https://www39085.034dapp.com/_layouts/15/spolanding.aspx", "https://www38024.068dapp.com/_layouts/15/spolanding.aspx", "https://www37045.007dapp.com/_layouts/15/spolanding.aspx", "https://www30090.054dapp.com/_layouts/15/spolanding.aspx", "https://www95027.027dapp.com/_layouts/15/spolanding.aspx", "https://www75007.023dapp.com/_layouts/15/spolanding.aspx", "https://www70030.035dapp.com/_layouts/15/spolanding.aspx", "https://www60140.098dspoapp.com/_layouts/15/spolanding.aspx", "https://www160015.078dapp.com/_layouts/15/spolanding.aspx", "https://www154017.003dapp.com/_layouts/15/spolanding.aspx", "https://www102027.067dapp.com/_layouts/15/spolanding.aspx", "https://www100039.017dapp.com/_layouts/15/spolanding.aspx", "https://www87072.042dapp.com/_layouts/15/spolanding.aspx", "https://www90082.053dapp.com/_layouts/15/spolanding.aspx", "https://www80033.011dapp.com/_layouts/15/spolanding.aspx", "https://www65158.013dspoapp.com/_layouts/15/spolanding.aspx", "https://www139017.073dapp.com/_layouts/15/spolanding.aspx", "https://www133018.046dapp.com/_layouts/15/spolanding.aspx", "https://www97058.085dspoapp.com/_layouts/15/spolanding.aspx"], "resourceSpecificApplicationPermissions": [], "samlSingleSignOnSettings": null, "servicePrincipalNames": ["00000003-0000-0ff1-ce00-000000000000/*.sharepoint.com", "00000003-0000-0ff1-ce00-000000000000", "https://microsoft.sharepoint-df.com"], "servicePrincipalType": "Application", "tags": [], "tokenEncryptionKeyId": null, "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "addIns": [], "appRoles": [{ "allowedMemberTypes": ["Application"], "description": "Allows the app to create, read, update, and delete documents and list items in all site collections without a signed in user.", "displayName": "Read and write items in all site collections", "id": "fbcd29d2-fcca-4405-aded-518d457caae4", "isEnabled": true, "origin": "Application", "value": "Sites.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read documents and list items in all site collections without a signed in user.", "displayName": "Read items in all site collections", "id": "d13f72ca-a275-4b96-b789-48ebcc4da984", "isEnabled": true, "origin": "Application", "value": "Sites.Read.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to have full control of all site collections without a signed in user.", "displayName": "Have full control of all site collections", "id": "678536fe-1083-478a-9c59-b99265e6b0d3", "isEnabled": true, "origin": "Application", "value": "Sites.FullControl.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read, create, update, and delete document libraries and lists in all site collections without a signed in user.", "displayName": "Read and write items and lists in all site collections", "id": "9bff6588-13f2-4c48-bbf2-ddab62256b36", "isEnabled": true, "origin": "Application", "value": "Sites.Manage.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read enterprise managed metadata and to read basic site info without a signed in user.", "displayName": "Read managed metadata", "id": "2a8d57a5-4090-4a41-bf1c-3c621d2ccad3", "isEnabled": true, "origin": "Application", "value": "TermStore.Read.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to write enterprise managed metadata and to read basic site info without a signed in user.", "displayName": "Read and write managed metadata", "id": "c8e3537c-ec53-43b9-bed3-b2bd3617ae97", "isEnabled": true, "origin": "Application", "value": "TermStore.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read and update user profiles and to read basic site info without a signed in user.", "displayName": "Read and write user profiles", "id": "741f803b-c850-494e-b5df-cde7c675a1ca", "isEnabled": true, "origin": "Application", "value": "User.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read user profiles without a signed in user.", "displayName": "Read user profiles", "id": "df021288-bdef-4463-88db-98f22de89214", "isEnabled": true, "origin": "Application", "value": "User.Read.All" }], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "oauth2PermissionScopes": [{ "adminConsentDescription": "Allows the app to read all OData reporting data from all ProjectWebApp site collections for the signed-in user.", "adminConsentDisplayName": "Read ProjectWebApp OData reporting data", "id": "a4c14cd7-8bd6-4337-8e87-78623dfc023b", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read all OData reporting data from all ProjectWebApp site collections for the signed-in user.", "userConsentDisplayName": "Read ProjectWebApp OData reporting data", "value": "ProjectWebAppReporting.Read" }, { "adminConsentDescription": "Allows the app to submit project task status updates the signed-in user.", "adminConsentDisplayName": "Submit project task status updates", "id": "c4258712-0efb-41f1-b6bc-be58e4e32f3f", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to submit project task status updates the signed-in user.", "userConsentDisplayName": "Submit project task status updates", "value": "TaskStatus.Submit" }, { "adminConsentDescription": "Allows the app to read, create, update, and delete the current user’s enterprise resources.", "adminConsentDisplayName": "Read and write user project enterprise resources", "id": "2511a087-5795-4cae-9123-d5b7d6ec4844", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read, create, update, and delete the current user’s enterprise resources.", "userConsentDisplayName": "Read and write user project enterprise resources", "value": "EnterpriseResource.Write" }, { "adminConsentDescription": "Allows the app to read the current user's enterprise resources.", "adminConsentDisplayName": "Read user project enterprise resources", "id": "b8341dab-4143-49da-8eb9-3d8c073f9e77", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read the current user's enterprise resources.", "userConsentDisplayName": "Read user project enterprise resources", "value": "EnterpriseResource.Read" }, { "adminConsentDescription": "Allows the app to read, create, update, and delete the current users’ projects.", "adminConsentDisplayName": "Read and write user projects", "id": "d75a7b17-f04e-40d9-8e35-79b949bdb891", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read, create, update, and delete the current users’ projects.", "userConsentDisplayName": "Read and write user projects", "value": "Project.Write" }, { "adminConsentDescription": "Allows the app to read the current user's projects.", "adminConsentDisplayName": "Read user projects", "id": "2beb830c-70d1-4f5b-a983-79cbdb0c6c6a", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read the current user's projects.", "userConsentDisplayName": "Read user projects", "value": "Project.Read" }, { "adminConsentDescription": "Allows the app to have full control of all ProjectWebApp site collections the signed-in user.", "adminConsentDisplayName": "Have full control of all ProjectWebApp site collections", "id": "e7e732bd-932b-45c4-8ce5-40d60a7daad9", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to have full control of all ProjectWebApp site collections the signed-in user.", "userConsentDisplayName": "Have full control of all ProjectWebApp site collections", "value": "ProjectWebApp.FullControl" }, { "adminConsentDescription": "Allows the app to read managed metadata and to read basic site info on behalf of the signed-in user.", "adminConsentDisplayName": "Read managed metadata", "id": "a468ea40-458c-4cc2-80c4-51781af71e41", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to read managed metadata and to read basic site info on your behalf.", "userConsentDisplayName": "Read managed metadata", "value": "TermStore.Read.All" }, { "adminConsentDescription": "Allows the app to read, create, update, and delete managed metadata and to read basic site info on behalf of the signed-in user.", "adminConsentDisplayName": "Read and write managed metadata", "id": "59a198b5-0420-45a8-ae59-6da1cb640505", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to read, create, update, and delete managed metadata and to read basic site info on your behalf.", "userConsentDisplayName": "Read and write managed metadata", "value": "TermStore.ReadWrite.All" }, { "adminConsentDescription": "Allows the app to run search queries and to read basic site info on behalf of the current signed-in user. Search results are based on the user's permissions instead of the app's permissions.", "adminConsentDisplayName": "Run search queries as a user", "id": "1002502a-9a71-4426-8551-69ab83452fab", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to run search queries and to read basic site info on your behalf. Search results are based on your permissions.", "userConsentDisplayName": "Run search queries ", "value": "Sites.Search.All" }, { "adminConsentDescription": "Allows the app to read documents and list items in all site collections on behalf of the signed-in user.", "adminConsentDisplayName": "Read items in all site collections", "id": "4e0d77b0-96ba-4398-af14-3baa780278f4", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read documents and list items in all site collections on your behalf.", "userConsentDisplayName": "Read items in all site collections", "value": "AllSites.Read" }, { "adminConsentDescription": "Allows the app to create, read, update, and delete documents and list items in all site collections on behalf of the signed-in user.", "adminConsentDisplayName": "Read and write items in all site collections", "id": "640ddd16-e5b7-4d71-9690-3f4022699ee7", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to create, read, update, and delete documents and list items in all site collections on your behalf.", "userConsentDisplayName": "Read and write items in all site collections", "value": "AllSites.Write" }, { "adminConsentDescription": "Allows the app to read, create, update, and delete document libraries and lists in all site collections on behalf of the signed-in user.", "adminConsentDisplayName": "Read and write items and lists in all site collections", "id": "b3f70a70-8a4b-4f95-9573-d71c496a53f4", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read, create, update, and delete document libraries and lists in all site collections on your behalf.", "userConsentDisplayName": "Read and write items and lists in all site collections", "value": "AllSites.Manage" }, { "adminConsentDescription": "Allows the app to have full control of all site collections on behalf of the signed-in user.", "adminConsentDisplayName": "Have full control of all site collections", "id": "56680e0d-d2a3-4ae1-80d8-3c4f2100e3d0", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to have full control of all site collections on your behalf.", "userConsentDisplayName": "Have full control of all site collections", "value": "AllSites.FullControl" }, { "adminConsentDescription": "Allows the app to read the current user's files.", "adminConsentDisplayName": "Read user files", "id": "dd2c8d78-58e1-46d7-82dd-34d411282686", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read your files.", "userConsentDisplayName": "Read your files", "value": "MyFiles.Read" }, { "adminConsentDescription": "Allows the app to read, create, update, and delete the current user's files.", "adminConsentDisplayName": "Read and write user files", "id": "2cfdc887-d7b4-4798-9b33-3d98d6b95dd2", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to read, create, update, and delete your files.", "userConsentDisplayName": "Read and write your files", "value": "MyFiles.Write" }, { "adminConsentDescription": "Allows the app to read and update user profiles and to read basic site info on behalf of the signed-in user.", "adminConsentDisplayName": "Read and write user profiles", "id": "82866913-39a9-4be7-8091-f4fa781088ae", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to read and update user profiles and to read basic site info on your behalf.", "userConsentDisplayName": "Read and write user profiles", "value": "User.ReadWrite.All" }, { "adminConsentDescription": "Allows the app to read user profiles and to read basic site info on behalf of the signed-in user.", "adminConsentDisplayName": "Read user profiles", "id": "0cea5a30-f6f8-42b5-87a0-84cc26822e02", "isEnabled": true, "type": "Admin", "userConsentDescription": "Allows the app to read user profiles and basic site info on your behalf.", "userConsentDisplayName": "Read user profiles", "value": "User.Read.All" }], "passwordCredentials": [] }] });
      }
      return Promise.reject();
    });
  };

  const postRequestStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve({ "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals('24448e9c-d0fa-43d1-a1dd-e279720969a0')/appRoleAssignments/$entity", "id": "nI5EJPrQ0UOh3eJ5cglpoLL3KmM12wZPom8Zw6AEypw", "deletedDateTime": null, "appRoleId": "9bff6588-13f2-4c48-bbf2-ddab62256b36", "createdDateTime": "2020-10-18T20:04:23.2456334Z", "principalDisplayName": "myapp", "principalId": "24448e9c-d0fa-43d1-a1dd-e279720969a0", "principalType": "ServicePrincipal", "resourceDisplayName": "Office 365 SharePoint Online", "resourceId": "df3d00f0-a24d-45a9-ba8b-3b0934ec3a6c" });
    });
  };

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
    loggerLogSpy = sinon.spy(logger, 'log');
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

    command.action(logger, { options: { displayName: 'myapp', resource: 'SharePoint', scope: 'Sites.Read.All' } }, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].objectId, 'nI5EJPrQ0UOh3eJ5cglpoLL3KmM12wZPom8Zw6AEypw');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].principalDisplayName, 'myapp');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].resourceDisplayName, 'Office 365 SharePoint Online');
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

    command.action(logger, { options: { objectId: '24448e9c-d0fa-43d1-a1dd-e279720969a0', resource: 'SharePoint', scope: 'Sites.Read.All,Sites.ReadWrite.All' } }, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].objectId, 'nI5EJPrQ0UOh3eJ5cglpoLL3KmM12wZPom8Zw6AEypw');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].principalDisplayName, 'myapp');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].resourceDisplayName, 'Office 365 SharePoint Online');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][1].objectId, 'nI5EJPrQ0UOh3eJ5cglpoLL3KmM12wZPom8Zw6AEypw');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][1].principalDisplayName, 'myapp');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][1].resourceDisplayName, 'Office 365 SharePoint Online');
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

    command.action(logger, { options: { displayName: 'myapp', resource: 'SharePoint', scope: 'Sites.Read.All', output: 'json' } }, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 'nI5EJPrQ0UOh3eJ5cglpoLL3KmM12wZPom8Zw6AEypw');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].principalDisplayName, 'myapp');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].resourceDisplayName, 'Office 365 SharePoint Online');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].principalId, '24448e9c-d0fa-43d1-a1dd-e279720969a0');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].principalType, 'ServicePrincipal');
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].resourceId, 'df3d00f0-a24d-45a9-ba8b-3b0934ec3a6c');
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

    command.action(logger, { options: { debug: true, appId: '26e49d05-4227-4ace-ae52-9b8f08f37184', resource: 'SharePoint', scope: 'Sites.Read.All' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
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

    command.action(logger, { options: { debug: true, appId: '26e49d05-4227-4ace-ae52-9b8f08f37184', resource: 'intune', scope: 'Sites.Read.All' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
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

    command.action(logger, { options: { debug: true, appId: '26e49d05-4227-4ace-ae52-9b8f08f37184', resource: 'exchange', scope: 'Sites.Read.All' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
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

    command.action(logger, { options: { debug: true, appId: '26e49d05-4227-4ace-ae52-9b8f08f37184', resource: 'fff194f1-7dce-4428-8301-1badb5518201', scope: 'Sites.Read.All' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('rejects if app roles are not found for the specified resource option value', (done) => {
    postRequestStub();
    sinon.stub(request, 'get').callsFake((opts: any): Promise<any> => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?`) > -1) {
        // fake first call for getting service principal
        if (opts.url.indexOf('startswith') === -1) {
          return Promise.resolve({ "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals", "value": [{ "id": "24448e9c-d0fa-43d1-a1dd-e279720969a0", "deletedDateTime": null, "accountEnabled": true, "alternativeNames": [], "appDisplayName": "myapp", "appDescription": null, "appId": "26e49d05-4227-4ace-ae52-9b8f08f37184", "applicationTemplateId": null, "appOwnerOrganizationId": "c8e571e1-d528-43d9-8776-dc51157d615a", "appRoleAssignmentRequired": false, "createdDateTime": "2020-08-29T18:35:13Z", "description": null, "displayName": "myapp", "homepage": null, "loginUrl": null, "logoutUrl": null, "notes": null, "notificationEmailAddresses": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyThumbprint": null, "replyUrls": ["https://login.microsoftonline.com/common/oauth2/nativeclient"], "resourceSpecificApplicationPermissions": [], "samlSingleSignOnSettings": null, "servicePrincipalNames": ["26e49d05-4227-4ace-ae52-9b8f08f37184"], "servicePrincipalType": "Application", "signInAudience": "AzureADMyOrg", "tags": ["WindowsAzureActiveDirectoryIntegratedApp"], "tokenEncryptionKeyId": null, "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "addIns": [], "appRoles": [], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "oauth2PermissionScopes": [], "passwordCredentials": [] }] });
        }
        // second get request for searching for service principals by resource options value specified
        return Promise.resolve({ "value": [{ objectId: "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "appRoles": [] }] });
      }
      return Promise.reject();
    });

    command.action(logger, { options: { debug: true, appId: '26e49d05-4227-4ace-ae52-9b8f08f37184', resource: 'SharePoint', scope: 'Sites.Read.All' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The resource 'SharePoint' does not have any application permissions available.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('rejects if app role scope not found for the specified resource option value', (done) => {
    postRequestStub();
    sinon.stub(request, 'get').callsFake((opts: any): Promise<any> => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?`) > -1) {
        // fake first call for getting service principal
        if (opts.url.indexOf('startswith') === -1) {
          return Promise.resolve({ "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals", "value": [{ "id": "24448e9c-d0fa-43d1-a1dd-e279720969a0", "deletedDateTime": null, "accountEnabled": true, "alternativeNames": [], "appDisplayName": "myapp", "appDescription": null, "appId": "26e49d05-4227-4ace-ae52-9b8f08f37184", "applicationTemplateId": null, "appOwnerOrganizationId": "c8e571e1-d528-43d9-8776-dc51157d615a", "appRoleAssignmentRequired": false, "createdDateTime": "2020-08-29T18:35:13Z", "description": null, "displayName": "myapp", "homepage": null, "loginUrl": null, "logoutUrl": null, "notes": null, "notificationEmailAddresses": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyThumbprint": null, "replyUrls": ["https://login.microsoftonline.com/common/oauth2/nativeclient"], "resourceSpecificApplicationPermissions": [], "samlSingleSignOnSettings": null, "servicePrincipalNames": ["26e49d05-4227-4ace-ae52-9b8f08f37184"], "servicePrincipalType": "Application", "signInAudience": "AzureADMyOrg", "tags": ["WindowsAzureActiveDirectoryIntegratedApp"], "tokenEncryptionKeyId": null, "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "addIns": [], "appRoles": [], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "oauth2PermissionScopes": [], "passwordCredentials": [] }] });
        }
        // second get request for searching for service principals by resource options value specified
        return Promise.resolve({ "value": [{ objectId: "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "appRoles": [{ value: 'Scope1', id: '1' }, { value: 'Scope2', id: '2' }] }] });
      }
      return Promise.reject();
    });

    command.action(logger, { options: { debug: true, appId: '26e49d05-4227-4ace-ae52-9b8f08f37184', resource: 'SharePoint', scope: 'Sites.Read.All' } } as any, (err?: any) => {
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
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?`) > -1) {
        // fake first call for getting service principal
        if (opts.url.indexOf('startswith') === -1) {
          return Promise.resolve({ "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals", "value": [{ "id": "24448e9c-d0fa-43d1-a1dd-e279720969a0", "deletedDateTime": null, "accountEnabled": true, "alternativeNames": [], "appDisplayName": "myapp", "appDescription": null, "appId": "26e49d05-4227-4ace-ae52-9b8f08f37184", "applicationTemplateId": null, "appOwnerOrganizationId": "c8e571e1-d528-43d9-8776-dc51157d615a", "appRoleAssignmentRequired": false, "createdDateTime": "2020-08-29T18:35:13Z", "description": null, "displayName": "myapp", "homepage": null, "loginUrl": null, "logoutUrl": null, "notes": null, "notificationEmailAddresses": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyThumbprint": null, "replyUrls": ["https://login.microsoftonline.com/common/oauth2/nativeclient"], "resourceSpecificApplicationPermissions": [], "samlSingleSignOnSettings": null, "servicePrincipalNames": ["26e49d05-4227-4ace-ae52-9b8f08f37184"], "servicePrincipalType": "Application", "signInAudience": "AzureADMyOrg", "tags": ["WindowsAzureActiveDirectoryIntegratedApp"], "tokenEncryptionKeyId": null, "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "addIns": [], "appRoles": [], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "oauth2PermissionScopes": [], "passwordCredentials": [] }, { "id": "24448e9c-d0fa-43d1-a1dd-e279720969a0", "deletedDateTime": null, "accountEnabled": true, "alternativeNames": [], "appDisplayName": "myapp", "appDescription": null, "appId": "26e49d05-4227-4ace-ae52-9b8f08f37184", "applicationTemplateId": null, "appOwnerOrganizationId": "c8e571e1-d528-43d9-8776-dc51157d615a", "appRoleAssignmentRequired": false, "createdDateTime": "2020-08-29T18:35:13Z", "description": null, "displayName": "myapp", "homepage": null, "loginUrl": null, "logoutUrl": null, "notes": null, "notificationEmailAddresses": [], "preferredSingleSignOnMode": null, "preferredTokenSigningKeyThumbprint": null, "replyUrls": ["https://login.microsoftonline.com/common/oauth2/nativeclient"], "resourceSpecificApplicationPermissions": [], "samlSingleSignOnSettings": null, "servicePrincipalNames": ["26e49d05-4227-4ace-ae52-9b8f08f37184"], "servicePrincipalType": "Application", "signInAudience": "AzureADMyOrg", "tags": ["WindowsAzureActiveDirectoryIntegratedApp"], "tokenEncryptionKeyId": null, "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "addIns": [], "appRoles": [], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "oauth2PermissionScopes": [], "passwordCredentials": [] }] });
        }
        // second get request for searching for service principals by resource options value specified
        return Promise.resolve({ "value": [{ objectId: "5edf62fd-ae7a-4a99-af2e-fc5950aaed07", "appRoles": [{ value: 'Scope1', id: '1' }, { value: 'Scope2', id: '2' }] }] });
      }
      return Promise.reject();
    });

    command.action(logger, { options: { debug: true, appId: '26e49d05-4227-4ace-ae52-9b8f08f37184', resource: 'SharePoint', scope: 'Sites.Read.All' } } as any, (err?: any) => {
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
    sinon.stub(request, 'get').callsFake(() => {
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

    command.action(logger, { options: { debug: false, appId: '36e3a540-6f25-4483-9542-9f5fa00bb633' } } as any, (err?: any) => {
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
    const actual = command.validate({ options: { resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = command.validate({ options: { appId: '123', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', () => {
    const actual = command.validate({ options: { objectId: '123', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and displayName are specified', () => {
    const actual = command.validate({ options: { appId: '123', displayName: 'abc', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both objectId and displayName are specified', () => {
    const actual = command.validate({ options: { objectId: '123', displayName: 'abc', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both objectId, appId and displayName are specified', () => {
    const actual = command.validate({ options: { appId: '123', objectId: '123', displayName: 'abc', resource: 'abc', scope: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appId option specified', () => {
    const actual = command.validate({ options: { appId: '57907bf8-73fa-43a6-89a5-1f603e29e452', resource: 'abc', scope: 'abc' } });
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

  it('supports specifying appId', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying displayName', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});

