import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./app-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.APP_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.APP_LIST), true);
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
        assert.equal(telemetry.name, commands.APP_LIST);
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

  it('retrieves available apps from the tenant app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if (opts.url.indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [
              {
                ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
                Title: 'online-client-side-solution',
                Deployed: true,
                AppCatalogVersion: '1.0.0.0'
              },
              {
                ID: 'e5f65aef-68fe-45b0-801e-92733dd57e2c',
                Title: 'onprem-client-side-solution',
                Deployed: true,
                AppCatalogVersion: '1.0.0.0'
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          },
          {
            ID: 'e5f65aef-68fe-45b0-801e-92733dd57e2c',
            Title: 'onprem-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('retrieves available apps from the site app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/sitecollectionappcatalog/AvailableApps') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [
              {
                ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
                Title: 'online-client-side-solution',
                Deployed: true,
                AppCatalogVersion: '1.0.0.0'
              },
              {
                ID: 'e5f65aef-68fe-45b0-801e-92733dd57e2c',
                Title: 'onprem-client-side-solution',
                Deployed: true,
                AppCatalogVersion: '1.0.0.0'
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          },
          {
            ID: 'e5f65aef-68fe-45b0-801e-92733dd57e2c',
            Title: 'onprem-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('includes all properties for output json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }

      if (opts.url.indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [
              {
                "AppCatalogVersion": "1.0.0.0",
                "CanUpgrade": false,
                "CurrentVersionDeployed": false,
                "Deployed": false,
                "ID": "b2307a39-e878-458b-bc90-03bc578531d6",
                "InstalledVersion": "",
                "IsClientSideSolution": true,
                "Title": "online-client-side-solution"
              },
              {
                "AppCatalogVersion": "1.0.0.0",
                "CanUpgrade": false,
                "CurrentVersionDeployed": false,
                "Deployed": false,
                "ID": "e6362993-d4fd-4c5a-8254-fd095a7291ad",
                "InstalledVersion": "",
                "IsClientSideSolution": true,
                "Title": "spfx-140-online-client-side-solution"
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "AppCatalogVersion": "1.0.0.0",
            "CanUpgrade": false,
            "CurrentVersionDeployed": false,
            "Deployed": false,
            "ID": "b2307a39-e878-458b-bc90-03bc578531d6",
            "InstalledVersion": "",
            "IsClientSideSolution": true,
            "Title": "online-client-side-solution"
          },
          {
            "AppCatalogVersion": "1.0.0.0",
            "CanUpgrade": false,
            "CurrentVersionDeployed": false,
            "Deployed": false,
            "ID": "e6362993-d4fd-4c5a-8254-fd095a7291ad",
            "InstalledVersion": "",
            "IsClientSideSolution": true,
            "Title": "spfx-140-online-client-side-solution"
          }
        ]))
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('correctly handles no apps in the tenant app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if (opts.url.indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify({ value: [] }));
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert.equal(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('handles if tenant appcatalog is null or not exist (debug)', (done) => {
    sinon.stub(request, 'get').resolves(JSON.stringify({ "CorporateCatalogUrl": null }));
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Tenant app catalog is not configured.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no apps in the site app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/sitecollectionappcatalog/AvailableApps') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify({ value: [] }));
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, () => {
      try {
        assert.equal(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('correctly handles no apps in the tenant app catalog (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if (opts.url.indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify({ value: [] }));
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, verbose: true } }, () => {
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('No apps found') > -1) {
          correctLogStatement = true;
        }
      })
      try {
        assert(correctLogStatement);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
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

  it('fails validation when invalid scope is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { scope: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when no scope is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.equal(actual, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', () => {
    const actual = (command.validate() as CommandValidate)({ options: { scope: 'tenant' } });
    assert.equal(actual, true);
  });

  it('fails validation when appCatalogUrl is not a valid url', () => {
    const actual = (command.validate() as CommandValidate)({ options: { scope: 'sitecollection', appCatalogUrl: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when appCatalogUrl and no scope', () => {
    const actual = (command.validate() as CommandValidate)({ options: { appCatalogUrl: 'https://contoso-admin.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', () => {

    const actual = (command.validate() as CommandValidate)({ options: { name: 'solution', filePath: 'abc', scope: 'sitecollection' } });
    assert.notEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but  bad appCatalogUrl format specified', () => {

    const actual = (command.validate() as CommandValidate)({ options: { name: 'solution', filePath: 'abc', scope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'sitecollection\' and appCatalogUrl present', () => {
    const actual = (command.validate() as CommandValidate)({ options: { scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } });
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
    assert(find.calledWith(commands.APP_LIST));
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
    sinon.stub(request, 'get').resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
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