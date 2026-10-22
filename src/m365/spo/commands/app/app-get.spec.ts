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
import command, { options } from './app-get.js';

describe(commands.APP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      request.get
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

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'solution.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com/sites/site1' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'solution.sppkg' }) });
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ name: 'solution.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' }) }),
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

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'solution.sppkg' }) });
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6' }) }),
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso-admin.sharepoint.com' }) }),
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6' }) }),
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6' }) }),
      new CommandError('An error has occurred'));
  });

  it('fails validation if neither the id nor the name options are specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `Specify either 'id' or 'name', but not both.`);
  });

  it('passes validation when the id option specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'f8b52a45-61d5-4264-81c9-c3bbd203e7d0' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the specified id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: '123' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, '123 is not a valid GUID');
  });

  it('fails validation when both the id and the name options specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'f8b52a45-61d5-4264-81c9-c3bbd203e7d0', name: 'solution.sppkg' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `Specify either 'id' or 'name', but not both.`);
  });

  it('fails validation when the specified appCatalogUrl is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'solution.sppkg', appCatalogUrl: 'url' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'url' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation when invalid scope is specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'foo' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Invalid option: expected one of "tenant"|"sitecollection"');
  });

  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogScope: 'sitecollection' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'You must specify appCatalogUrl when the appCatalogScope is sitecollection');
  });

  it('should pass when scope \'tenant\' and appCatalogUrl specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogScope: 'tenant', appCatalogUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });

  it('should fail when \'sitecollection\' scope, but  bad appCatalogUrl format specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogScope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'contoso.sharepoint.com' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation when the name option specified', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'solution.sppkg' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when the name and appCatalogUrl options specified', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps', appCatalogScope: 'tenant' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when no scope is specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogScope: 'tenant' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when the scope is specified with \'sitecollection\'', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', appCatalogUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'sitecollection' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'dd20afdf-d7fd-4662-a443-b69e65a72bd4', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Unrecognized key: "unknownOption"');
  });
});
