import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./app-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.APP_GET, () => {
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
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.APP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the app with the specified id from the tenant app catalog (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
          Title: 'online-client-side-solution',
          Deployed: true,
          AppCatalogVersion: '1.0.0.0'
        }));
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

  it('retrieves information about the app with the specified id from the site app catalog (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
          Title: 'online-client-side-solution',
          Deployed: true,
          AppCatalogVersion: '1.0.0.0'
        }));
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

  it('retrieves information about the app with the specified id from the tenant app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
          Title: 'online-client-side-solution',
          Deployed: true,
          AppCatalogVersion: '1.0.0.0'
        }));
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

  it('retrieves information about the app with the specified id from the site app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6', scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
          Title: 'online-client-side-solution',
          Deployed: true,
          AppCatalogVersion: '1.0.0.0'
        }));
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

  it('retrieves information about the app with the specified name from the specified tenant app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          });
        }
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return Promise.resolve({
          UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
          Title: 'online-client-side-solution',
          Deployed: true,
          AppCatalogVersion: '1.0.0.0'
        }));
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

  it('retrieves information about the app with the specified name from the specified site app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          });
        }
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/site1/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return Promise.resolve({
          UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'solution.sppkg', scope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com/sites/site1' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
          Title: 'online-client-side-solution',
          Deployed: true,
          AppCatalogVersion: '1.0.0.0'
        }));
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

  it('retrieves information about the app with the specified name with auto-discovered tenant app catalog (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          });
        }
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return Promise.resolve({
          UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: 'solution.sppkg' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
          Title: 'online-client-side-solution',
          Deployed: true,
          AppCatalogVersion: '1.0.0.0'
        }));
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

  it('should handle getfolderbyserverrelativeurl error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return Promise.reject({
          error: {
            'odata.error': {
              code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
              message: {
                value: 'An error has occurred'
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { appCatalogUrl: string, }) => void) => {
      cb({ appCatalogUrl: '' });
    };
    cmdInstance.action({ options: { debug: false, name: 'solution.sppkg', scope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the app with the specified name from the specified tenant app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return Promise.resolve({
          UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { appCatalogUrl: string }) => void) => {
      cb({ appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps' });
    };
    cmdInstance.action({ options: { debug: false, name: 'solution.sppkg' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
          Title: 'online-client-side-solution',
          Deployed: true,
          AppCatalogVersion: '1.0.0.0'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no app found in the tenant app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.reject({
            error: {
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                message: {
                  lang: "en-US",
                  value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                }
              }
            }
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
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

  it('correctly handles no app found in the site app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.reject({
            error: {
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                message: {
                  lang: "en-US",
                  value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                }
              }
            }
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6', scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
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

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return Promise.resolve({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.reject({
            error: {
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
                message: {
                  value: 'An error has occurred'
                }
              }
            }
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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

  it('fails validation if neither the id nor the name options are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid scope is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '123', appCatalogUrl: 'https://contoso.sharepoint.com', scope: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'f8b52a45-61d5-4264-81c9-c3bbd203e7d0' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the specified id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both the id and the name options specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'f8b52a45-61d5-4264-81c9-c3bbd203e7d0', name: 'solution.sppkg' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified appCatalogUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'solution.sppkg', appCatalogUrl: 'url' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid scope is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogUrl: 'https://contoso.sharepoint.com', scope: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', filePath: 'abc', scope: 'sitecollection' } });
    assert.notStrictEqual(actual, true);
  });

  it('should pass when scope \'tenant\' and appCatalogUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', filePath: 'abc', scope: 'tenant', appCatalogUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but  bad appCatalogUrl format specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', filePath: 'abc', scope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'f8b52a45-61d5-4264-81c9-c3bbd203e7d0' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the name option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'solution.sppkg' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the name and appCatalogUrl options specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps', scope: 'tenant' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when no scope is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', scope: 'tenant' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'sitecollection\'', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogUrl: 'https://contoso.sharepoint.com', scope: 'sitecollection' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when no scope is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});