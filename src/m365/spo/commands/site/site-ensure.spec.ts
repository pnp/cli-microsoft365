import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import { WebProperties } from '../web/WebProperties.js';
import command, { options } from './site-ensure.js';

describe(commands.SITE_ENSURE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const webResponse: WebProperties = {
    AllowRssFeeds: true,
    AlternateCssUrl: '',
    AppInstanceId: '00000000-0000-0000-0000-000000000000',
    Configuration: 0,
    Created: '2021-01-22T18:39:51.06',
    CurrentChangeToken: {
      StringValue: '1;2;113ba5b6-c737-4a6b-b1c7-2a367290057e;637470248884630000;125942029'
    },
    CustomMasterUrl: '/sites/team1/_catalogs/masterpage/seattle.master',
    Description: 'Team 2',
    DesignPackageId: '00000000-0000-0000-0000-000000000000',
    DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
    EnableMinimalDownload: false,
    HorizontalQuickLaunch: false,
    Id: '113ba5b6-c737-4a6b-b1c7-2a367290057e',
    IsMultilingual: true,
    Language: 1033,
    LastItemModifiedDate: '2021-01-22T18:44:16Z',
    LastItemUserModifiedDate: '2021-01-22T18:39:57Z',
    MasterUrl: '/sites/team1/_catalogs/masterpage/seattle.master',
    NoCrawl: false,
    OverwriteTranslationsOnChange: false,
    ResourcePath: {
      DecodedUrl: 'https://contoso.sharepoint.com/sites/team1'
    },
    QuickLaunchEnabled: true,
    RecycleBinEnabled: true,
    ServerRelativeUrl: '/sites/team1',
    SiteLogoUrl: '',
    SyndicationEnabled: true,
    Title: 'Team 2 updated',
    TreeViewEnabled: false,
    UIVersion: 15,
    UIVersionConfigurationEnabled: false,
    Url: 'https://contoso.sharepoint.com/sites/team1',
    WebTemplate: 'GROUP',
    AssociatedMemberGroup: '',
    AssociatedOwnerGroup: '',
    AssociatedVisitorGroup: ''
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      spo.getWeb,
      spo.addSite,
      spo.updateSite
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_ENSURE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates modern team site if no site found', async () => {
    sinon.stub(spo, 'getWeb').rejects({ error: '404 FILE NOT FOUND' });
    sinon.stub(spo, 'addSite').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1' }) });
  });

  it('creates modern communication site if no site found (debug)', async () => {
    sinon.stub(spo, 'getWeb').rejects({ error: '404 FILE NOT FOUND' });
    sinon.stub(spo, 'addSite').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/comms', title: 'Comms', type: 'CommunicationSite', debug: true }) });
  });

  it('updates modern team site if existing modern team site found (no type specified)', async () => {
    sinon.stub(spo, 'getWeb').resolves(webResponse);
    sinon.stub(spo, 'updateSite').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1' }) });
  });

  it('updates modern team site if existing modern team site found (type specified)', async () => {
    sinon.stub(spo, 'getWeb').resolves(webResponse);
    sinon.stub(spo, 'updateSite').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1', type: 'TeamSite' }) });
  });

  it('updates modern communication site if existing modern communication site found (no type specified; debug)', async () => {
    const webResponseSitePagePublishing = { ...webResponse };
    webResponseSitePagePublishing.WebTemplate = 'SITEPAGEPUBLISHING';

    sinon.stub(spo, 'getWeb').resolves(webResponseSitePagePublishing);
    sinon.stub(spo, 'updateSite').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/commsite1', title: 'CommSite1', debug: true }) });
  });

  it('updates modern communication site if existing modern communication site found (type specified)', async () => {
    const webResponseSitePagePublishing = { ...webResponse };
    webResponseSitePagePublishing.WebTemplate = 'SITEPAGEPUBLISHING';

    sinon.stub(spo, 'getWeb').resolves(webResponseSitePagePublishing);
    sinon.stub(spo, 'updateSite').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/commsite1', title: 'CommSite1', type: 'CommunicationSite' }) });
  });

  it('updates classic site if an existing classic site found (type specified)', async () => {
    const webResponseSts = { ...webResponse };
    webResponseSts.WebTemplate = 'STS';

    sinon.stub(spo, 'getWeb').resolves(webResponseSts);
    sinon.stub(spo, 'updateSite').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/classic', title: 'Classic', type: 'ClassicSite' }) });
  });

  it(`updates site's visibility and sharing options`, async () => {
    sinon.stub(spo, 'getWeb').resolves(webResponse);
    sinon.stub(spo, 'updateSite').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1', isPublic: true, shareByEmailEnabled: true }) });
  });

  it('returns error when validation of options for creating site failed', async () => {
    sinon.stub(spo, 'getWeb').rejects(new Error('404 FILE NOT FOUND'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1' }) }));
  });

  it('returns error when an error has occurred when checking if a site exists at the specified URL', async () => {
    sinon.stub(spo, 'getWeb').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1' }) }));
  });

  // it('returns error when the specified site type is invalid', async () => {
  //   const webResponseSts = { ...webResponse };
  //   webResponseSts.WebTemplate = 'STS';

  //   sinon.stub(spo, 'getWeb').resolves(webResponseSts);

  //   await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/classic', title: 'Classic', type: 'Invalid' }) }));
  // });

  it('returns error when a communication site expected but a team site found', async () => {
    sinon.stub(spo, 'getWeb').resolves(webResponse);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1', type: 'CommunicationSite' }) }));
  });

  it('fails validation if url is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'contoso.sharepoint.com', title: 'Team 1' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'contoso.sharepoint.com' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if type is not a valid site type', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1', type: 'Invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if siteDesign is not a valid value', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1', siteDesign: 'Invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if sharingCapability is not a valid value', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1', sharingCapability: 'Invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if title is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/team1' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if url is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ title: 'Team 1' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});
