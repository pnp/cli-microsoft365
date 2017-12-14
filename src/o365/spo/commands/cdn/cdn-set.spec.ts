import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./cdn-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.CDN_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
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
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    assert.equal(command.name.startsWith(commands.CDN_SET), true);
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
        assert.equal(telemetry.name, commands.CDN_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint tenant admin site', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(`https://contoso.sharepoint.com is not a tenant admin site. Connect to your tenant admin site and try again`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('enables public CDN when Public type specified and enabled set to true', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, enabled: 'true', type: 'Public' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('disables public CDN when Public type specified and enabled set to false', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, enabled: 'false', type: 'Public' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables private CDN when Private type specified and enabled set to true', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, enabled: 'true', type: 'Private' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('enables public CDN when no type specified and enabled set to true', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, enabled: 'true' } }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers['X-RequestDigest'] &&
          r.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

  it('correctly handles an error when setting tenant CDN settings', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
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
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, enabled: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
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

  it('requires tenant enabled state', () => {
    const options = (command.options() as CommandOption[]);
    let requiresOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<enabled>') > -1) {
        requiresOption = true;
      }
    });
    assert(requiresOption);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return undefined; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('accepts Public SharePoint Online CDN type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { type: 'Public', enabled: 'true' } });
    assert(actual);
  });

  it('accepts Private SharePoint Online CDN type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { type: 'Private', enabled: 'true' } });
    assert(actual);
  });

  it('rejects invalid SharePoint Online CDN type', () => {
    const type = 'foo';
    const actual = (command.validate() as CommandValidate)({ options: { type: type, enabled: 'true' } });
    assert.equal(actual, `${type} is not a valid CDN type. Allowed values are Public|Private`);
  });

  it('doesn\'t fail validation if the optional type option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { enabled: 'true' } });
    assert(actual);
  });

  it('accepts true SharePoint Online CDN enabled state', () => {
    const actual = (command.validate() as CommandValidate)({ options: { enabled: 'true' } });
    assert(actual);
  });

  it('accepts True SharePoint Online CDN enabled state', () => {
    const actual = (command.validate() as CommandValidate)({ options: { enabled: 'True' } });
    assert(actual);
  });

  it('accepts TRUE SharePoint Online CDN enabled state', () => {
    const actual = (command.validate() as CommandValidate)({ options: { enabled: 'TRUE' } });
    assert(actual);
  });

  it('accepts false SharePoint Online CDN enabled state', () => {
    const actual = (command.validate() as CommandValidate)({ options: { enabled: 'false' } });
    assert(actual);
  });

  it('accepts False SharePoint Online CDN enabled state', () => {
    const actual = (command.validate() as CommandValidate)({ options: { enabled: 'False' } });
    assert(actual);
  });

  it('accepts FALSE SharePoint Online CDN enabled state', () => {
    const actual = (command.validate() as CommandValidate)({ options: { enabled: 'FALSE' } });
    assert(actual);
  });

  it('rejects invalid SharePoint Online CDN enabled state', () => {
    const enabled = 'foo';
    const actual = (command.validate() as CommandValidate)({ options: { enabled: enabled } });
    assert.equal(actual, `${enabled} is not a valid boolean value. Allowed values are true|false`);
  });

  it('fails validation if the required enabled state option is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { } });
    assert.equal(actual, 'undefined is not a valid boolean value. Allowed values are true|false');
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
    assert(find.calledWith(commands.CDN_SET));
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
    cmdInstance.action({ options: { debug: true, confirm: true, key: 'existingproperty' }, appCatalogUrl: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});