import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-list.js';

describe(commands.APP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), [`Title`, `ID`, `Deployed`, `AppCatalogVersion`]);
  });

  it('retrieves available apps from the tenant app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
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
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([
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
    ]));
  });

  it('retrieves available apps from the site app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitecollectionappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
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
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } });
    assert(loggerLogSpy.calledWith([
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
    ]));
  });

  it('includes all properties for output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }

      if ((opts.url as string).indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
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
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, output: 'json' } });
    assert(loggerLogSpy.calledWith([
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
    ]));
  });

  it('correctly handles no apps in the tenant app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf('/_api/web/tenantappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { value: [] };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('handles if tenant appcatalog is null or not exist (debug)', async () => {
    sinon.stub(request, 'get').resolves(JSON.stringify({ "CorporateCatalogUrl": null }));
    await assert.rejects(command.action(logger, {
      options: {
        debug: true
      }
    } as any), new CommandError('Tenant app catalog is not configured.'));
  });

  it('correctly handles no apps in the site app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitecollectionappcatalog/AvailableApps') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { value: [] };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('fails validation when invalid scope is specified', async () => {
    const actual = await command.validate({ options: { appCatalogScope: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when no scope is specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', async () => {
    const actual = await command.validate({ options: { appCatalogScope: 'tenant' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when appCatalogUrl is not a valid url', async () => {
    const actual = await command.validate({ options: { appCatalogScope: 'sitecollection', appCatalogUrl: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', async () => {
    const actual = await command.validate({ options: { appCatalogScope: 'sitecollection' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but  bad appCatalogUrl format specified', async () => {
    const actual = await command.validate({ options: { appCatalogScope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'sitecollection\' and appCatalogUrl present', async () => {
    const actual = await command.validate({ options: { appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
