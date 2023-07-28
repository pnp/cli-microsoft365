import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-appcatalog-add.js';
import { spo } from '../../../../utils/spo.js';

describe(commands.TENANT_APPCATALOG_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const siteResponse = {
    AllowCreateDeclarativeWorkflow: true,
    AllowDesigner: true,
    AllowMasterPageEditing: false,
    AllowRevertFromTemplate: false,
    AllowSaveDeclarativeWorkflowAsTemplate: true,
    AllowSavePublishDeclarativeWorkflow: true,
    AllowSelfServiceUpgrade: true,
    AllowSelfServiceUpgradeEvaluation: true,
    AuditLogTrimmingRetention: 90,
    Classification: '',
    CompatibilityLevel: 15,
    CurrentChangeToken: {
      StringValue: '1;1;1a70e568-d286-4ad1-b036-734ff8667915;636527399616270000;66855110'
    },
    DisableAppViews: false,
    DisableCompanyWideSharingLinks: false,
    DisableFlows: false,
    ExternalSharingTipsEnabled: false,
    GeoLocation: 'EUR',
    GroupId: '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091',
    HubSiteId: '00000000-0000-0000-0000-000000000000',
    Id: '1a70e568-d286-4ad1-b036-734ff8667915',
    IsHubSite: false,
    LockIssue: null,
    MaxItemsPerThrottledOperation: 5000,
    NeedsB2BUpgrade: false,
    ResourcePath: {
      DecodedUrl: 'https://contoso.sharepoint.com/sites/appCatalog'
    },
    PrimaryUri: 'https://contoso.sharepoint.com/sites/appCatalog',
    ReadOnly: false,
    RequiredDesignerVersion: '15.0.0.0',
    SandboxedCodeActivationCapability: 2,
    ServerRelativeUrl: '/sites/appCatalog',
    ShareByEmailEnabled: true,
    ShareByLinkEnabled: false,
    ShowUrlStructure: false,
    TrimAuditLog: true,
    UIVersionConfigurationEnabled: false,
    UpgradeReminderDate: '1899-12-30T00:00:00',
    UpgradeScheduled: false,
    UpgradeScheduledDate: '1753-01-01T00:00:00',
    Upgrading: false,
    Url: 'https://contoso.sharepoint.com/sites/appCatalog'
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      spo.removeSite,
      spo.addSite,
      spo.getSite,
      spo.getTenantAppCatalogUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_APPCATALOG_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any);
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used (debug)', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any);
  });

  it('handles error when creating app catalog when app catalog and site with different URL already exist and force used failed', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when app catalog and site with different URL already exist, force used and deleting the existing site failed', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').rejects(new Error('Error deleting site new-app-catalog'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('Error deleting site new-app-catalog'));
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').callsFake(async (url) => {
      if (url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
        return siteResponse;
      }

      if (url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
        throw new Error('404 FILE NOT FOUND');
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any);
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used (debug)', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').callsFake(async (url) => {
      if (url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
        return siteResponse;
      }

      if (url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
        throw new Error('404 FILE NOT FOUND');
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any);
  });

  it('handles error when creating app catalog when app catalog already exists, site with different URL does not exist and force used', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').callsFake(async (url) => {
      if (url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
        return siteResponse;
      }

      if (url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
        throw new Error('404 FILE NOT FOUND');
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when retrieving site with different URL failed and app catalog already exists, and force used', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').callsFake(async (url) => {
      if (url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
        return siteResponse;
      }

      if (url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
        throw new Error('An error has occurred');
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when deleting existing app catalog failed', async () => {
    sinon.stub(spo, 'removeSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error app catalog exists and no force used', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/old-app-catalog'));
  });

  it('creates app catalog when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').callsFake(async (url) => {
      if (url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
        throw new Error('404 FILE NOT FOUND');
      }

      if (url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
        return siteResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any);
  });

  it('handles error when creating app catalog when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').callsFake(async (url) => {
      if (url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
        throw new Error('404 FILE NOT FOUND');
      }

      if (url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
        return siteResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when deleting existing site, when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(spo, 'removeSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').callsFake(async (url) => {
      if (url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
        throw new Error('404 FILE NOT FOUND');
      }

      if (url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
        return siteResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when app catalog does not exist, site with different URL already exists and force not used', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').callsFake(async (url) => {
      if (url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
        throw new Error('404 FILE NOT FOUND');
      }

      if (url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
        return siteResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog'));
  });

  it(`creates app catalog when app catalog and site with different URL don't exist`, async () => {
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').rejects(new Error('404 FILE NOT FOUND'));

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any);
  });

  it(`handles error when creating app catalog fails, when app catalog when app catalog does and site with different URL don't exist`, async () => {
    sinon.stub(spo, 'addSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').rejects(new Error('404 FILE NOT FOUND'));

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when checking if the app catalog site exists`, async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/old-app-catalog');
    sinon.stub(spo, 'getSite').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any);
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used (debug)`, async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any);
  });

  it(`handles error when creating app catalog when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(spo, 'removeSite').resolves();
    sinon.stub(spo, 'addSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when deleting existing site when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(spo, 'removeSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when app catalog not registered, site with different URL exists and force not used`, async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('');
    sinon.stub(spo, 'getSite').resolves(siteResponse);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog'));
  });

  it(`creates app catalog when app catalog not registered and site with different URL doesn't exist`, async () => {
    sinon.stub(spo, 'addSite').resolves();
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('');
    sinon.stub(spo, 'getSite').rejects(new Error('404 FILE NOT FOUND'));

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any);
  });

  it(`handles error when creating app catalog when app catalog not registered and site with different URL doesn't exist`, async () => {
    sinon.stub(spo, 'addSite').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('');
    sinon.stub(spo, 'getSite').rejects(new Error('404 FILE NOT FOUND'));

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when app catalog not registered and checking if the site with different URL exists throws error`, async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('');
    sinon.stub(spo, 'getSite').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when checking if app catalog registered throws error`, async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if the specified url is not a valid SharePoint URL', async () => {
    const options: any = { url: '/foo', owner: 'user@contoso.com', timeZone: 4 };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(typeof actual, 'string');
  });

  it('fails validation if timeZone is not a number', async () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 'a' };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(typeof actual, 'string');
  });

  it('passes validation when all options are specified and valid', async () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 4 };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
