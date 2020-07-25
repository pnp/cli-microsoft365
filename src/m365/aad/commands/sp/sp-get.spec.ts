import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./sp-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SP_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    assert.strictEqual(command.name.startsWith(commands.SP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified service principal using its display name (debug)', (done) => {
    const sp: any = { "objectType": "ServicePrincipal", "objectId": "d03a0062-1aa6-43e1-8f49-d73e969c5812", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": [], "appDisplayName": "SharePoint Online Client", "appId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "SharePoint Online Client", "errorUrl": null, "homepage": null, "keyCredentials": [], "logoutUrl": null, "oauth2Permissions": [], "passwordCredentials": [], "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "servicePrincipalNames": ["57fb890c-0dab-4253-a5e0-7188c88b2bb4"], "servicePrincipalType": "Application", "tags": [], "tokenEncryptionKeyId": null };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6&$filter=displayName eq 'SharePoint%20Online%20Client'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ value: [sp] });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, displayName: 'SharePoint Online Client' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(sp));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified service principal using its display name', (done) => {
    const sp: any = { "objectType": "ServicePrincipal", "objectId": "d03a0062-1aa6-43e1-8f49-d73e969c5812", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": [], "appDisplayName": "SharePoint Online Client", "appId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "SharePoint Online Client", "errorUrl": null, "homepage": null, "keyCredentials": [], "logoutUrl": null, "oauth2Permissions": [], "passwordCredentials": [], "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "servicePrincipalNames": ["57fb890c-0dab-4253-a5e0-7188c88b2bb4"], "servicePrincipalType": "Application", "tags": [], "tokenEncryptionKeyId": null };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6&$filter=displayName eq 'SharePoint%20Online%20Client'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ value: [sp] });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, displayName: 'SharePoint Online Client' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(sp));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified service principal using its appId', (done) => {
    const sp: any = { "objectType": "ServicePrincipal", "objectId": "d03a0062-1aa6-43e1-8f49-d73e969c5812", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": [], "appDisplayName": "SharePoint Online Client", "appId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "SharePoint Online Client", "errorUrl": null, "homepage": null, "keyCredentials": [], "logoutUrl": null, "oauth2Permissions": [], "passwordCredentials": [], "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "servicePrincipalNames": ["57fb890c-0dab-4253-a5e0-7188c88b2bb4"], "servicePrincipalType": "Application", "tags": [], "tokenEncryptionKeyId": null };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6&$filter=appId eq '57fb890c-0dab-4253-a5e0-7188c88b2bb4'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ value: [sp] });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, appId: '57fb890c-0dab-4253-a5e0-7188c88b2bb4' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(sp));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified service principal using its objectId', (done) => {
    const sp: any = { "objectType": "ServicePrincipal", "objectId": "d03a0062-1aa6-43e1-8f49-d73e969c5812", "deletionTimestamp": null, "accountEnabled": true, "addIns": [], "alternativeNames": [], "appDisplayName": "SharePoint Online Client", "appId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "appOwnerTenantId": null, "appRoleAssignmentRequired": false, "appRoles": [], "displayName": "SharePoint Online Client", "errorUrl": null, "homepage": null, "keyCredentials": [], "logoutUrl": null, "oauth2Permissions": [], "passwordCredentials": [], "preferredTokenSigningKeyThumbprint": null, "publisherName": null, "replyUrls": [], "samlMetadataUrl": null, "servicePrincipalNames": ["57fb890c-0dab-4253-a5e0-7188c88b2bb4"], "servicePrincipalType": "Application", "tags": [], "tokenEncryptionKeyId": null };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6&$filter=objectId eq '57fb890c-0dab-4253-a5e0-7188c88b2bb4'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ value: [sp] });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, objectId: '57fb890c-0dab-4253-a5e0-7188c88b2bb4' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(sp));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no service principal found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6&$filter=displayName eq 'Foo'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ value: [] });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, displayName: 'Foo' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither the appId nor the displayName option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appId option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '6a7b1395-d313-4682-8ed4-65a6265a6320' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the displayName option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'Microsoft Graph' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when both the appId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '6a7b1395-d313-4682-8ed4-65a6265a6320', displayName: 'Microsoft Graph' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { objectId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appId: '123', displayName: 'abc' } });
    assert.notStrictEqual(actual, true);
  })

  it('fails validation if objectId and displayName are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'abc', objectId: '123' } });
    assert.notStrictEqual(actual, true);
  })

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