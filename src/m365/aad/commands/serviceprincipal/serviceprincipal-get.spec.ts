import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./serviceprincipal-get');

describe(commands.SERVICEPRINCIPAL_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.SERVICEPRINCIPAL_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if either id, appId and name options are not passed', (done) => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both id and name options are passed', (done) => {
    const actual = command.validate({
      options: {
        id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
        name: 'Service Principal app'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both appId and name options are passed', (done) => {
    const actual = command.validate({
      options: {
        appId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
        name: 'Service Principal app'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both id and appId options are passed', (done) => {
    const actual = command.validate({
      options: {
        id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
        appId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if all id, appId and name options are passed', (done) => {
    const actual = command.validate({
      options: {
        id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
        appId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
        name: 'Service Principal app'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = command.validate({ options: { id: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = command.validate({ options: { appId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', () => {
    const actual = command.validate({ options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the appId is a valid GUID', () => {
    const actual = command.validate({ options: { appId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert.strictEqual(actual, true);
  });

  it('fails to get service principal information due to wrong service principal id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/serviceprincipals/1caf7dcd-7e83-4c3a-94f7-932a1299c843`) {
        return Promise.reject({
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c843' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2021-09-26T03:05:45",
              "request-id": "236e9ef1-9be4-43d1-8602-64e5550e23bd",
              "client-request-id": "562eff66-8944-b2a2-dbd1-11ad0618b539"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c843' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, "Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c843' does not exist or one of its queried reference-property objects are not present.");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get service principal when name does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/serviceprincipals?$filter=displayName eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject(`The specified service principal doesn't exist in Azure AD`);
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'Service Principal app'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified service principal doesn't exist in Azure AD`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get  when appId does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/serviceprincipals?$filter=appId eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('The specified  does not exist in the Azure AD directory');
    });

    command.action(logger, {
      options: {
        debug: true,
        appId: '8a2a376d-5f57-4c14-9639-692f841c00bc'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified service principal doesn't exist in Azure AD`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when service principal name does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/serviceprincipals?$filter=displayName eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject(`The specified service principal doesn't exist in Azure AD`);
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'Service Principal app'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified service principal doesn't exist in Azure AD`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple service principals with same name exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/serviceprincipals?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 2,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "00000000-0000-0000-0000-000000000000"
            }
          ]
        }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'Service Principal app'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple service principals with name Service Principal app found: 00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple service principals with same appId exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/serviceprincipals?$filter=appId eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 2,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "00000000-0000-0000-0000-000000000000"
            }
          ]
        }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        appId: '8a2a376d-5f57-4c14-9639-692f841c00bd'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple service principals with appId 8a2a376d-5f57-4c14-9639-692f841c00bd found: 00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified service principal', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/serviceprincipals/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {

        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "accountEnabled": true,
          "alternativeNames": [],
          "appDisplayName": "Service Principal App",
          "appDescription": null,
          "appId": "8a2a376d-5f57-4c14-9639-692f841c00bd",
          "applicationTemplateId": null,
          "appOwnerOrganizationId": "00000000-0000-0000-0000-000000000000",
          "appRoleAssignmentRequired": false,
          "createdDateTime": "2018-10-05T22:54:37Z",
          "description": null,
          "disabledByMicrosoftStatus": null,
          "displayName": "Service Principal App",
          "homepage": null,
          "loginUrl": null,
          "logoutUrl": null,
          "notes": null,
          "notificationEmailAddresses": [],
          "preferredSingleSignOnMode": null,
          "preferredTokenSigningKeyThumbprint": null,
          "replyUrls": [
            "http://localhost:8080/"
          ],
          "servicePrincipalNames": [
            "8a2a376d-5f57-4c14-9639-692f841c00bd"
          ],
          "servicePrincipalType": "Application",
          "signInAudience": "AzureADandPersonalMicrosoftAccount",
          "tags": [
            "WindowsAzureActiveDirectoryIntegratedApp"
          ],
          "tokenEncryptionKeyId": null,
          "resourceSpecificApplicationPermissions": [],
          "samlSingleSignOnSettings": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "addIns": [],
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "oauth2PermissionScopes": [],
          "passwordCredentials": []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "accountEnabled": true,
          "alternativeNames": [],
          "appDisplayName": "Service Principal App",
          "appDescription": null,
          "appId": "8a2a376d-5f57-4c14-9639-692f841c00bd",
          "applicationTemplateId": null,
          "appOwnerOrganizationId": "00000000-0000-0000-0000-000000000000",
          "appRoleAssignmentRequired": false,
          "createdDateTime": "2018-10-05T22:54:37Z",
          "description": null,
          "disabledByMicrosoftStatus": null,
          "displayName": "Service Principal App",
          "homepage": null,
          "loginUrl": null,
          "logoutUrl": null,
          "notes": null,
          "notificationEmailAddresses": [],
          "preferredSingleSignOnMode": null,
          "preferredTokenSigningKeyThumbprint": null,
          "replyUrls": [
            "http://localhost:8080/"
          ],
          "servicePrincipalNames": [
            "8a2a376d-5f57-4c14-9639-692f841c00bd"
          ],
          "servicePrincipalType": "Application",
          "signInAudience": "AzureADandPersonalMicrosoftAccount",
          "tags": [
            "WindowsAzureActiveDirectoryIntegratedApp"
          ],
          "tokenEncryptionKeyId": null,
          "resourceSpecificApplicationPermissions": [],
          "samlSingleSignOnSettings": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "addIns": [],
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "oauth2PermissionScopes": [],
          "passwordCredentials": []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified service principal by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/v1.0/serviceprincipals?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/serviceprincipals/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {

        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "accountEnabled": true,
          "alternativeNames": [],
          "appDisplayName": "Service Principal App",
          "appDescription": null,
          "appId": "8a2a376d-5f57-4c14-9639-692f841c00bd",
          "applicationTemplateId": null,
          "appOwnerOrganizationId": "00000000-0000-0000-0000-000000000000",
          "appRoleAssignmentRequired": false,
          "createdDateTime": "2018-10-05T22:54:37Z",
          "description": null,
          "disabledByMicrosoftStatus": null,
          "displayName": "Service Principal App",
          "homepage": null,
          "loginUrl": null,
          "logoutUrl": null,
          "notes": null,
          "notificationEmailAddresses": [],
          "preferredSingleSignOnMode": null,
          "preferredTokenSigningKeyThumbprint": null,
          "replyUrls": [
            "http://localhost:8080/"
          ],
          "servicePrincipalNames": [
            "8a2a376d-5f57-4c14-9639-692f841c00bd"
          ],
          "servicePrincipalType": "Application",
          "signInAudience": "AzureADandPersonalMicrosoftAccount",
          "tags": [
            "WindowsAzureActiveDirectoryIntegratedApp"
          ],
          "tokenEncryptionKeyId": null,
          "resourceSpecificApplicationPermissions": [],
          "samlSingleSignOnSettings": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "addIns": [],
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "oauth2PermissionScopes": [],
          "passwordCredentials": []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: 'Service Principal App' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "accountEnabled": true,
          "alternativeNames": [],
          "appDisplayName": "Service Principal App",
          "appDescription": null,
          "appId": "8a2a376d-5f57-4c14-9639-692f841c00bd",
          "applicationTemplateId": null,
          "appOwnerOrganizationId": "00000000-0000-0000-0000-000000000000",
          "appRoleAssignmentRequired": false,
          "createdDateTime": "2018-10-05T22:54:37Z",
          "description": null,
          "disabledByMicrosoftStatus": null,
          "displayName": "Service Principal App",
          "homepage": null,
          "loginUrl": null,
          "logoutUrl": null,
          "notes": null,
          "notificationEmailAddresses": [],
          "preferredSingleSignOnMode": null,
          "preferredTokenSigningKeyThumbprint": null,
          "replyUrls": [
            "http://localhost:8080/"
          ],
          "servicePrincipalNames": [
            "8a2a376d-5f57-4c14-9639-692f841c00bd"
          ],
          "servicePrincipalType": "Application",
          "signInAudience": "AzureADandPersonalMicrosoftAccount",
          "tags": [
            "WindowsAzureActiveDirectoryIntegratedApp"
          ],
          "tokenEncryptionKeyId": null,
          "resourceSpecificApplicationPermissions": [],
          "samlSingleSignOnSettings": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "addIns": [],
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "oauth2PermissionScopes": [],
          "passwordCredentials": []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('retrieves information about the specified service principal by appId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/v1.0/serviceprincipals?$filter=appId eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/serviceprincipals/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {

        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "accountEnabled": true,
          "alternativeNames": [],
          "appDisplayName": "Service Principal App",
          "appDescription": null,
          "appId": "8a2a376d-5f57-4c14-9639-692f841c00bd",
          "applicationTemplateId": null,
          "appOwnerOrganizationId": "00000000-0000-0000-0000-000000000000",
          "appRoleAssignmentRequired": false,
          "createdDateTime": "2018-10-05T22:54:37Z",
          "description": null,
          "disabledByMicrosoftStatus": null,
          "displayName": "Service Principal App",
          "homepage": null,
          "loginUrl": null,
          "logoutUrl": null,
          "notes": null,
          "notificationEmailAddresses": [],
          "preferredSingleSignOnMode": null,
          "preferredTokenSigningKeyThumbprint": null,
          "replyUrls": [
            "http://localhost:8080/"
          ],
          "servicePrincipalNames": [
            "8a2a376d-5f57-4c14-9639-692f841c00bd"
          ],
          "servicePrincipalType": "Application",
          "signInAudience": "AzureADandPersonalMicrosoftAccount",
          "tags": [
            "WindowsAzureActiveDirectoryIntegratedApp"
          ],
          "tokenEncryptionKeyId": null,
          "resourceSpecificApplicationPermissions": [],
          "samlSingleSignOnSettings": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "addIns": [],
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "oauth2PermissionScopes": [],
          "passwordCredentials": []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, appId: '8a2a376d-5f57-4c14-9639-692f841c00bd' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "accountEnabled": true,
          "alternativeNames": [],
          "appDisplayName": "Service Principal App",
          "appDescription": null,
          "appId": "8a2a376d-5f57-4c14-9639-692f841c00bd",
          "applicationTemplateId": null,
          "appOwnerOrganizationId": "00000000-0000-0000-0000-000000000000",
          "appRoleAssignmentRequired": false,
          "createdDateTime": "2018-10-05T22:54:37Z",
          "description": null,
          "disabledByMicrosoftStatus": null,
          "displayName": "Service Principal App",
          "homepage": null,
          "loginUrl": null,
          "logoutUrl": null,
          "notes": null,
          "notificationEmailAddresses": [],
          "preferredSingleSignOnMode": null,
          "preferredTokenSigningKeyThumbprint": null,
          "replyUrls": [
            "http://localhost:8080/"
          ],
          "servicePrincipalNames": [
            "8a2a376d-5f57-4c14-9639-692f841c00bd"
          ],
          "servicePrincipalType": "Application",
          "signInAudience": "AzureADandPersonalMicrosoftAccount",
          "tags": [
            "WindowsAzureActiveDirectoryIntegratedApp"
          ],
          "tokenEncryptionKeyId": null,
          "resourceSpecificApplicationPermissions": [],
          "samlSingleSignOnSettings": null,
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "addIns": [],
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "oauth2PermissionScopes": [],
          "passwordCredentials": []
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