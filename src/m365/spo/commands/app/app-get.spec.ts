import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./app-get');

describe(commands.APP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6', scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: false, name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: false, name: 'solution.sppkg', scope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com/sites/site1' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: true, name: 'solution.sppkg' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        sinonUtil.restore(request.get);
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

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { appCatalogUrl: string, }) => void) => {
      cb({ appCatalogUrl: '' });
    });
    command.action(logger, { options: { debug: false, name: 'solution.sppkg', scope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the app with the specified name from the specified tenant app catalog via prompt', (done) => {
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { appCatalogUrl: string }) => void) => {
      cb({ appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps' });
    });
    command.action(logger, { options: { debug: false, name: 'solution.sppkg' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6', scope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore(request.get);
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
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore(request.get);
      }
    });
  });

  it('fails validation if neither the id nor the name options are specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id option specified', async () => {
    const actual = await command.validate({ options: { id: 'f8b52a45-61d5-4264-81c9-c3bbd203e7d0' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the specified id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both the id and the name options specified', async () => {
    const actual = await command.validate({ options: { id: 'f8b52a45-61d5-4264-81c9-c3bbd203e7d0', name: 'solution.sppkg' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified appCatalogUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { name: 'solution.sppkg', appCatalogUrl: 'url' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid scope is specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogUrl: 'https://contoso.sharepoint.com', scope: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', scope: 'sitecollection' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should pass when scope \'tenant\' and appCatalogUrl specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', scope: 'tenant', appCatalogUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but  bad appCatalogUrl format specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', scope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name option specified', async () => {
    const actual = await command.validate({ options: { name: 'solution.sppkg' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the name and appCatalogUrl options specified', async () => {
    const actual = await command.validate({ options: { name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps', scope: 'tenant' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when no scope is specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', scope: 'tenant' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'sitecollection\'', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogUrl: 'https://contoso.sharepoint.com', scope: 'sitecollection' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});