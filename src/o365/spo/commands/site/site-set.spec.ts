import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./site-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.SITE_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      (command as any).getRequestDigestForSite
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SITE_SET), true);
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
        assert.equal(telemetry.name, commands.SITE_SET);
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Connect to a SharePoint Online site first')));
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`https://contoso.sharepoint.com is not a tenant admin site. Connect to your tenant admin site and try again`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification of the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_vti_bin/client.svc/ProcessQuery` &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
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
    cmdInstance.action({ options: { debug: false, classification: 'HBI', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('resets the classification of the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_vti_bin/client.svc/ProcessQuery` &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String"></Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
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
    cmdInstance.action({ options: { debug: false, classification: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets disableFlows to true for the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_vti_bin/client.svc/ProcessQuery` &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="28" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">true</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
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
    cmdInstance.action({ options: { debug: false, disableFlows: 'true', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets disableFlows to false for the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_vti_bin/client.svc/ProcessQuery` &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="28" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">false</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
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
    cmdInstance.action({ options: { debug: false, disableFlows: 'false', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification and disableFlows of the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_vti_bin/client.svc/ProcessQuery` &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty><SetProperty Id="28" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">true</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
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
    cmdInstance.action({ options: { debug: false, classification: 'HBI', disableFlows: 'true', id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification of the specified site. Retrieves the ID of the site collection (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_api/site?$select=Id`) {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_vti_bin/client.svc/ProcessQuery` &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
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
    cmdInstance.action({ options: { debug: true, classification: 'HBI', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the classification of the specified site. Retrieves the ID of the site collection', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_api/site?$select=Id`) {
        return Promise.resolve({
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_vti_bin/client.svc/ProcessQuery` &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:255a50b2-527f-4413-8485-57f4c17a24d1" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7317.1205", "ErrorInfo": null, "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
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
    cmdInstance.action({ options: { debug: false, classification: 'HBI', url: 'https://contoso.sharepoint.com/sites/Sales' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles site not found error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/Sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1', classification: 'HBI' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("404 - \"404 FILE NOT FOUND\"")));
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
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": {
              "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPException"
            }, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799"
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
    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/Sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1', classification: 'HBI' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('configures command types', () => {
    assert.notEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures classification as string option', () => {
    const types = (command.types() as CommandTypes);
    ['classification'].forEach(o => {
      assert.notEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
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

  it('supports specifying site ID', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying site url', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying site classification', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--classification') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying disableFlows', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--disableFlows') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if URL not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { classification: 'HBI' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if URL is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'Invalid', classification: 'HBI' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if specified id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com', id: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if no property to update specified (id not specified)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if no property to update specified (id specified)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if invalid value specified for disableFlows', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com', disableFlows: 'Invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if url and classification specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com', classification: 'HBI' } });
    assert.equal(actual, true);
  });

  it('passes validation if url and empty classification specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com', classification: '' } });
    assert.equal(actual, true);
  });

  it('passes validation if url and disableFlows true specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com', disableFlows: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation if url and disableFlows false specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url: 'https://contoso.sharepoint.com', disableFlows: 'false' } });
    assert.equal(actual, true);
  });

  it('passes validation if url, id, classification and disableFlows specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', url: 'https://contoso.sharepoint.com', classification: 'HBI', disableFlows: 'true' } });
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
    assert(find.calledWith(commands.SITE_SET));
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
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com' } }, (err?: any) => {
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