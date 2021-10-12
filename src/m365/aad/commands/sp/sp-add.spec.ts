import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./sp-add');

describe(commands.SP_ADD, () => {
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
    (command as any).items = [];
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
    assert.strictEqual(command.name.startsWith(commands.SP_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if neither the appId, appName, nor objectId option specified', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = command.validate({ options: { appId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', () => {
    const actual = command.validate({ options: { objectId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and appName are specified', () => {
    const actual = command.validate({ options: { appId: '00000000-0000-0000-0000-000000000000', appName: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appName and objectId are specified', () => {
    const actual = command.validate({ options: { appName: 'abc', objectId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and objectId are specified', () => {
    const actual = command.validate({ options: { appId: '00000000-0000-0000-0000-000000000000', objectId: '00000000-0000-0000-0000-000000000000' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appId option specified', () => {
    const actual = command.validate({ options: { appId: '00000000-0000-0000-0000-000000000000' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the appName option specified', () => {
    const actual = command.validate({ options: { appName: 'abc' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the objectId option specified', () => {
    const actual = command.validate({ options: { objectId: '00000000-0000-0000-0000-000000000000' } });
    assert.strictEqual(actual, true);
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

  it('supports specifying appName', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying objectId', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--objectId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no service principal found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ value: [] });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, appName: 'Foo' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when Azure AD app with same name exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=displayName eq `) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals",
          "value": [
            {
              "id": "be559819-b036-470f-858b-281c4e808403",
              "deletedDateTime": null,
              "accountEnabled": true,
              "alternativeNames": [],
              "appDisplayName": null,
              "appDescription": null,
              "appId": "ee091f63-9e48-4697-8462-7cfbf7410b8e",
              "applicationTemplateId": null,
              "appOwnerOrganizationId": null,
              "appRoleAssignmentRequired": false,
              "createdDateTime": "2021-03-07T15:04:11Z",
              "description": null,
              "disabledByMicrosoftStatus": null,
              "displayName": "Test App",
              "homepage": null,
              "loginUrl": null,
              "logoutUrl": null,
              "notes": null,
              "notificationEmailAddresses": [],
              "preferredSingleSignOnMode": null,
              "preferredTokenSigningKeyThumbprint": null,
              "replyUrls": [
                "https://localhost/"
              ],
              "servicePrincipalNames": [
                "ee091f63-9e48-4697-8462-7cfbf7410b8e",
                "ee091f63-9e48-4697-8462-7cfbf7410b8e/localhost"
              ],
              "servicePrincipalType": "Legacy",
              "signInAudience": null,
              "tags": [],
              "tokenEncryptionKeyId": null,
              "info": null,
              "resourceSpecificApplicationPermissions": [],
              "samlSingleSignOnSettings": null,
              "verifiedPublisher": {
                "displayName": null,
                "verifiedPublisherId": null,
                "addedDateTime": null
              },
              "addIns": [],
              "appRoles": [],
              "keyCredentials": [
                {
                  "customKeyIdentifier": null,
                  "displayName": null,
                  "endDateTime": "2022-03-07T15:04:11.0580414Z",
                  "key": null,
                  "keyId": "c535f925-afd5-498b-9e72-ee4c12822755",
                  "startDateTime": "2021-03-07T15:04:11.0580414Z",
                  "type": "Symmetric",
                  "usage": "Verify"
                },
                {
                  "customKeyIdentifier": null,
                  "displayName": null,
                  "endDateTime": "2022-03-07T15:04:11.0580414Z",
                  "key": null,
                  "keyId": "2bba7eac-b190-4df5-bcd6-78e9b3404be5",
                  "startDateTime": "2021-03-07T15:04:11.0580414Z",
                  "type": "Symmetric",
                  "usage": "Sign"
                }
              ],
              "oauth2PermissionScopes": [],
              "passwordCredentials": [
                {
                  "customKeyIdentifier": null,
                  "displayName": null,
                  "endDateTime": "2022-03-07T15:04:11.0580414Z",
                  "hint": null,
                  "keyId": "7aa974b3-083a-40b2-92db-cae026ab4482",
                  "secretText": null,
                  "startDateTime": "2021-03-07T15:04:11.0580414Z"
                }
              ]
            },
            {
              "id": "93d75ef9-ba9b-4361-9a47-1f6f7478f05f",
              "deletedDateTime": null,
              "accountEnabled": true,
              "alternativeNames": [],
              "appDisplayName": "Test App",
              "appDescription": null,
              "appId": "e9fd0957-049f-40d0-8d1d-112320fb1cbd",
              "applicationTemplateId": null,
              "appOwnerOrganizationId": "de348bc7-1aeb-4406-8cb3-97db021cadb4",
              "appRoleAssignmentRequired": false,
              "createdDateTime": "2021-10-12T10:41:23Z",
              "description": null,
              "disabledByMicrosoftStatus": null,
              "displayName": "Test App",
              "homepage": null,
              "loginUrl": null,
              "logoutUrl": null,
              "notes": null,
              "notificationEmailAddresses": [],
              "preferredSingleSignOnMode": null,
              "preferredTokenSigningKeyThumbprint": null,
              "replyUrls": [],
              "servicePrincipalNames": [
                "e9fd0957-049f-40d0-8d1d-112320fb1cbd"
              ],
              "servicePrincipalType": "Application",
              "signInAudience": "AzureADMyOrg",
              "tags": [
                "HideApp",
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
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        appName: 'Test App'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple Azure AD apps with name Test App found: ee091f63-9e48-4697-8462-7cfbf7410b8e,e9fd0957-049f-40d0-8d1d-112320fb1cbd`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a service principal to a registered Azure AD app by appId', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals`) {
        return Promise.resolve({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        appId: '65415bb1-9267-4313-bbf5-ae259732ee12'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a service principal to a registered Azure AD app by appName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=displayName eq `) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
              "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
              "displayName": "foo",
              "createdDateTime": "2021-03-07T15:04:11Z",
              "description": null,
              "homepage": null,
              "loginUrl": null,
              "logoutUrl": null,
              "notes": null
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        appName: 'foo'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a service principal to a registered Azure AD app by objectId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=id eq `) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
              "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
              "displayName": "foo",
              "createdDateTime": "2021-03-07T15:04:11Z",
              "description": null,
              "homepage": null,
              "loginUrl": null,
              "logoutUrl": null,
              "notes": null
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        objectId: '59e617e5-e447-4adc-8b88-00af644d7c92'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
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
