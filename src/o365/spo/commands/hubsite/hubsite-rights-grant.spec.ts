import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./hubsite-rights-grant');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.HUBSITE_RIGHTS_GRANT, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigestForSite').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
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
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth,
      (command as any).getRequestDigestForSite
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.HUBSITE_RIGHTS_GRANT), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
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
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.HUBSITE_RIGHTS_GRANT);
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`https://contoso.sharepoint.com is not a tenant admin site. Log in to your tenant admin site and try again`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principal', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principal (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principals', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object><Object Type="String">user</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin,user', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principals (email)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin@contoso.onmicrosoft.com</Object><Object Type="String">user@contoso.onmicrosoft.com</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin@contoso.onmicrosoft.com,user@contoso.onmicrosoft.com', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants rights on the specified site design to the specified principals separated with an extra space', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object><Object Type="String">user</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin, user', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in user input', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sales&gt;</Parameter><Parameter Type="Array"><Object Type="String">admin&gt;</Object></Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "1fbd439e-5090-5000-c29b-037f60060808"
          }, 37, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales>', principals: 'admin>', rights: 'Join' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": {
              "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "86be439e-80c4-5000-fcf8-b746bccdc4e7", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
            }, "TraceCorrelationId": "86be439e-80c4-5000-fcf8-b746bccdc4e7"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/sales', principals: 'admin', rights: 'Join' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports specifying hub site url', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying principals', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--principals') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying rights', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--rights') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if url not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { principals: 'admin', rights: 'Join' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if url is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'abc', principals: 'admin', rights: 'Join' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if principals not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', rights: 'Join' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if rights not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', principals: 'PattiF' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if specified rights value is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', principals: 'PattiF', rights: 'Invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', principals: 'PattiF', rights: 'Join' } });
    assert.equal(actual, true);
  });

  it('passes validation if all required parameters are valid (multiple principals)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com/sites/sales', principals: 'PattiF,AdeleV', rights: 'Join' } });
    assert.equal(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.HUBSITE_RIGHTS_GRANT));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
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