import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./storageentity-set');

describe(commands.STORAGEENTITY_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: any) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.STORAGEENTITY_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets tenant property', (done) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String">ipsum</Parameter><Parameter Type="String">dolor</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, key: 'Property1', value: 'Lorem', description: 'ipsum', comment: 'dolor', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual((postStub.lastCall.args[0].headers as any)['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String">ipsum</Parameter><Parameter Type="String">dolor</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets tenant property without description and comment', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, key: 'Property1', value: 'Lorem', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String"></Parameter><Parameter Type="String"></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets tenant property without description and comment (debug)', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, key: 'Property1', value: 'Lorem', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String"></Parameter><Parameter Type="String"></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in user input', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, key: '<Property1>', value: '\"Lorem\"', description: '"ipsum"', comment: '<dolor & samet>', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">&lt;Property1&gt;</Parameter><Parameter Type="String">&quot;Lorem&quot;</Parameter><Parameter Type="String">&quot;ipsum&quot;</Parameter><Parameter Type="String">&lt;dolor &amp; samet&gt;</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles a generic error when setting tenant property', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
            }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, key: 'Property1', value: 'Lorem', description: 'ipsum', comment: 'dolor', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } } as any, (err?: any) => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String">ipsum</Parameter><Parameter Type="String">dolor</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles access denied error when setting tenant property', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "Access denied.", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
            }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, key: 'Property1', value: 'Lorem', description: 'ipsum', comment: 'dolor', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } } as any, (err?: any) => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String">ipsum</Parameter><Parameter Type="String">dolor</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Access denied.')));
        assert.strictEqual(loggerSpy.calledWithMatch('This error is often caused by invalid URL of the app catalog site'), true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('requires app catalog URL', () => {
    const options = command.options();
    let requiresAppCatalogUrl = false;
    options.forEach(o => {
      if (o.option.indexOf('<appCatalogUrl>') > -1) {
        requiresAppCatalogUrl = true;
      }
    });
    assert(requiresAppCatalogUrl);
  });

  it('requires tenant property name', () => {
    const options = command.options();
    let requiresTenantPropertyName = false;
    options.forEach(o => {
      if (o.option.indexOf('<key>') > -1) {
        requiresTenantPropertyName = true;
      }
    });
    assert(requiresTenantPropertyName);
  });

  it('requires tenant property value', () => {
    const options = command.options();
    let requiresTenantPropertyValue = false;
    options.forEach(o => {
      if (o.option.indexOf('<value>') > -1) {
        requiresTenantPropertyValue = true;
      }
    });
    assert(requiresTenantPropertyValue);
  });

  it('supports setting tenant property description', () => {
    const options = command.options();
    let supportsTenantPropertyDescription = false;
    options.forEach(o => {
      if (o.option.indexOf('[description]') > -1) {
        supportsTenantPropertyDescription = true;
      }
    });
    assert(supportsTenantPropertyDescription);
  });

  it('supports setting tenant property comment', () => {
    const options = command.options();
    let supportsTenantPropertyComment = false;
    options.forEach(o => {
      if (o.option.indexOf('[comment]') > -1) {
        supportsTenantPropertyComment = true;
      }
    });
    assert(supportsTenantPropertyComment);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = command.options();
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('accepts valid SharePoint Online app catalog URL', () => {
    const actual = command.validate({ options: { appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert.strictEqual(actual, true);
  });

  it('accepts valid SharePoint Online site URL', () => {
    const actual = command.validate({ options: { appCatalogUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online URL', () => {
    const url = 'http://contoso';
    const actual = command.validate({ options: { appCatalogUrl: url } });
    assert.strictEqual(actual, `${url} is not a valid SharePoint Online site URL`);
  });

  it('fails validation when no SharePoint Online app catalog URL specified', () => {
    const actual = command.validate({ options: {} });
    assert.strictEqual(actual, 'Missing required option appCatalogUrl');
  });

  it('handles promise rejection', (done) => {
    sinon.stub(command as any, 'getSpoAdminUrl').callsFake(() => Promise.reject('error'));

    command.action(logger, {
      options: { debug: true, key: 'Property1', value: 'Lorem', description: 'ipsum', comment: 'dolor', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});