import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./storageentity-set');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.STORAGEENTITY_SET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String">ipsum</Parameter><Parameter Type="String">dolor</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: any) => {
        log.push(msg);
      }
    };
    auth.site = new Site();
    telemetry = null;
    requests = [];
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth,
      request.post
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.STORAGEENTITY_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {}, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {}, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert.equal(telemetry.name, commands.STORAGEENTITY_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint tenant admin site', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`${auth.site.url} is not a tenant admin site. Log in to your tenant admin site and try again`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets tenant property', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, key: 'Property1', value: 'Lorem', description: 'ipsum', comment: 'dolor', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String">ipsum</Parameter><Parameter Type="String">dolor</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`) {
          setRequestIssued = true;
        }
      });

      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets tenant property without description and comment', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String"></Parameter><Parameter Type="String"></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, key: 'Property1', value: 'Lorem', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.equal(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets tenant property without description and comment (debug)', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String"></Parameter><Parameter Type="String"></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, key: 'Property1', value: 'Lorem', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      let isDone = false;
      log.forEach(l => {
        if (l && typeof l === 'string' && l.indexOf('DONE')) {
          isDone = true;
        }
      });

      try {
        assert(isDone);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in user input', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">&lt;Property1&gt;</Parameter><Parameter Type="String">&apos;Lorem&apos;</Parameter><Parameter Type="String">&quot;ipsum&quot;</Parameter><Parameter Type="String">&lt;dolor &amp; samet&gt;</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, key: '<Property1>', value: '\'Lorem\'', description: '"ipsum"', comment: '<dolor & samet>', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      let isDone = false;
      log.forEach(l => {
        if (l && typeof l === 'string' && l.indexOf('DONE')) {
          isDone = true;
        }
      });

      try {
        assert(isDone);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles a generic error when setting tenant property', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String">ipsum</Parameter><Parameter Type="String">dolor</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                  "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
                }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
              }
            ]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, key: 'Property1', value: 'Lorem', description: 'ipsum', comment: 'dolor', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('correctly handles access denied error when setting tenant property', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.body) {
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">Property1</Parameter><Parameter Type="String">Lorem</Parameter><Parameter Type="String">ipsum</Parameter><Parameter Type="String">dolor</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([
              {
                "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                  "ErrorMessage": "Access denied.", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
                }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
              }
            ]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, key: 'Property1', value: 'Lorem', description: 'ipsum', comment: 'dolor', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      let accessDeniedErrorHandled = false;
      log.forEach(l => {
        if (l && typeof l === 'string' && l.indexOf('This error is often caused by invalid URL of the app catalog site.') > -1) {
          accessDeniedErrorHandled = true;
        }
      });

      try {
        assert(accessDeniedErrorHandled);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('requires app catalog URL', () => {
    const options = (command.options() as CommandOption[]);
    let requiresAppCatalogUrl = false;
    options.forEach(o => {
      if (o.option.indexOf('<appCatalogUrl>') > -1) {
        requiresAppCatalogUrl = true;
      }
    });
    assert(requiresAppCatalogUrl);
  });

  it('requires tenant property name', () => {
    const options = (command.options() as CommandOption[]);
    let requiresTenantPropertyName = false;
    options.forEach(o => {
      if (o.option.indexOf('<key>') > -1) {
        requiresTenantPropertyName = true;
      }
    });
    assert(requiresTenantPropertyName);
  });

  it('requires tenant property value', () => {
    const options = (command.options() as CommandOption[]);
    let requiresTenantPropertyValue = false;
    options.forEach(o => {
      if (o.option.indexOf('<value>') > -1) {
        requiresTenantPropertyValue = true;
      }
    });
    assert(requiresTenantPropertyValue);
  });

  it('supports setting tenant property description', () => {
    const options = (command.options() as CommandOption[]);
    let supportsTenantPropertyDescription = false;
    options.forEach(o => {
      if (o.option.indexOf('[description]') > -1) {
        supportsTenantPropertyDescription = true;
      }
    });
    assert(supportsTenantPropertyDescription);
  });

  it('supports setting tenant property comment', () => {
    const options = (command.options() as CommandOption[]);
    let supportsTenantPropertyComment = false;
    options.forEach(o => {
      if (o.option.indexOf('[comment]') > -1) {
        supportsTenantPropertyComment = true;
      }
    });
    assert(supportsTenantPropertyComment);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return undefined; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('accepts valid SharePoint Online app catalog URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert.equal(actual, true);
  });

  it('accepts valid SharePoint Online site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appCatalogUrl: 'https://contoso.sharepoint.com' } });
    assert.equal(actual, true);
  });

  it('rejects invalid SharePoint Online URL', () => {
    const url = 'http://contoso';
    const actual = (command.validate() as CommandValidate)({ options: { appCatalogUrl: url } });
    assert.equal(actual, `${url} is not a valid SharePoint Online site URL`);
  });

  it('fails validation when no SharePoint Online app catalog URL specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.equal(actual, 'Missing required option appCatalogUrl');
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => {},
      prompt: () => {},
      helpInformation: () => {}
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
    assert(find.calledWith(commands.STORAGEENTITY_SET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => {},
      helpInformation: () => {}
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, confirm: true, key: 'existingproperty', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});