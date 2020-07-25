import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./app-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.APP_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves available apps from the tenant app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
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
      if ((opts.url as string).indexOf('/_api/web/sitecollectionappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
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
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }

      if ((opts.url as string).indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
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
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify({ value: [] }));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert.strictEqual(log.length, 0);
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
    cmdInstance.action({
      options: {
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Tenant app catalog is not configured.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no apps in the site app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitecollectionappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify({ value: [] }));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, () => {
      try {
        assert.strictEqual(log.length, 0);
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
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify({ value: [] }));
        }
      }

      return Promise.reject('Invalid request');
    });

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
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when no scope is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', () => {
    const actual = (command.validate() as CommandValidate)({ options: { scope: 'tenant' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when appCatalogUrl is not a valid url', () => {
    const actual = (command.validate() as CommandValidate)({ options: { scope: 'sitecollection', appCatalogUrl: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'solution', filePath: 'abc', scope: 'sitecollection' } });
    assert.notStrictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but  bad appCatalogUrl format specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'solution', filePath: 'abc', scope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'sitecollection\' and appCatalogUrl present', () => {
    const actual = (command.validate() as CommandValidate)({ options: { scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
});