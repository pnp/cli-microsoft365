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
import command from './app-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_GET, () => {
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
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the app with the specified id from the tenant app catalog (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
    assert(loggerLogSpy.calledWith({
      ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
      Title: 'online-client-side-solution',
      Deployed: true,
      AppCatalogVersion: '1.0.0.0'
    }));
  });

  it('retrieves information about the app with the specified id from the site app catalog (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } });
    assert(loggerLogSpy.calledWith({
      ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
      Title: 'online-client-side-solution',
      Deployed: true,
      AppCatalogVersion: '1.0.0.0'
    }));
  });

  it('retrieves information about the app with the specified id from the tenant app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
    assert(loggerLogSpy.calledWith({
      ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
      Title: 'online-client-side-solution',
      Deployed: true,
      AppCatalogVersion: '1.0.0.0'
    }));
  });

  it('retrieves information about the app with the specified id from the site app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } });
    assert(loggerLogSpy.calledWith({
      ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
      Title: 'online-client-side-solution',
      Deployed: true,
      AppCatalogVersion: '1.0.0.0'
    }));
  });

  it('retrieves information about the app with the specified name from the specified tenant app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          };
        }
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return { UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6' };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps' } });
    assert(loggerLogSpy.calledWith({
      ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
      Title: 'online-client-side-solution',
      Deployed: true,
      AppCatalogVersion: '1.0.0.0'
    }));
  });

  it('retrieves information about the app with the specified name from the specified site app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          };
        }
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/site1/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return { UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6' };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'solution.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com/sites/site1' } });
    assert(loggerLogSpy.calledWith({
      ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
      Title: 'online-client-side-solution',
      Deployed: true,
      AppCatalogVersion: '1.0.0.0'
    }));
  });

  it('retrieves information about the app with the specified name with auto-discovered tenant app catalog (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          };
        }
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return { UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6' };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, name: 'solution.sppkg' } });
    assert(loggerLogSpy.calledWith({
      ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
      Title: 'online-client-side-solution',
      Deployed: true,
      AppCatalogVersion: '1.0.0.0'
    }));
  });

  it('should handle getfolderbyserverrelativeurl error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        throw {
          error: {
            'odata.error': {
              code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
              message: {
                value: 'An error has occurred'
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: 'solution.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
      new CommandError('An error has occurred'));
  });

  it('retrieves information about the app with the specified name from the specified tenant app catalog via prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) {
        return { UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6' };
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution',
            Deployed: true,
            AppCatalogVersion: '1.0.0.0'
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'solution.sppkg' } });
    assert(loggerLogSpy.calledWith({
      ID: 'b2307a39-e878-458b-bc90-03bc578531d6',
      Title: 'online-client-side-solution',
      Deployed: true,
      AppCatalogVersion: '1.0.0.0'
    }));
  });

  it('correctly handles no app found in the tenant app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw {
            error: {
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                message: {
                  lang: "en-US",
                  value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                }
              }
            }
          };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any),
      new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
  });

  it('correctly handles no app found in the site app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw {
            error: {
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                message: {
                  lang: "en-US",
                  value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                }
              }
            }
          };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' } } as any),
      new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw { error: 'An error has occurred' };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw {
            error: {
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
                message: {
                  value: 'An error has occurred'
                }
              }
            }
          };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if neither the id nor the name options are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: 'f8b52a45-61d5-4264-81c9-c3bbd203e7d0', name: 'solution.sppkg' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified appCatalogUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { name: 'solution.sppkg', appCatalogUrl: 'url' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid scope is specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogScope: 'sitecollection' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should pass when scope \'tenant\' and appCatalogUrl specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogScope: 'tenant', appCatalogUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but  bad appCatalogUrl format specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogScope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name option specified', async () => {
    const actual = await command.validate({ options: { name: 'solution.sppkg' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the name and appCatalogUrl options specified', async () => {
    const actual = await command.validate({ options: { name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps', appCatalogScope: 'tenant' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when no scope is specified', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogScope: 'tenant' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'sitecollection\'', async () => {
    const actual = await command.validate({ options: { id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'sitecollection' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
