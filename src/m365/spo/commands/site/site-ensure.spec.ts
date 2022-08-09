import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
import * as spoWebGetCommand from '../web/web-get';
import * as spoSiteAddCommand from './site-add';
import * as spoSiteSetCommand from './site-set';
const command: Command = require('./site-ensure');

describe(commands.SITE_ENSURE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.executeCommand,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_ENSURE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates modern team site if no site found', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.reject({
          error: new CommandError('404 FILE NOT FOUND')
        });
      }

      if (command === spoSiteAddCommand) {
        if (JSON.stringify(args) === JSON.stringify({
          options: {
            title: 'Team 1',
            alias: 'team1',
            verbose: false,
            debug: false,
            _: []
          }
        })) {
          return Promise.resolve({
            stdout: ''
          });
        }
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1' } } as any, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates modern communication site if no site found (debug)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.reject({
          error: new CommandError('404 FILE NOT FOUND')
        });
      }

      if (command === spoSiteAddCommand) {
        if (JSON.stringify(args) === JSON.stringify({
          options: {
            type: 'CommunicationSite',
            title: 'Comms',
            url: 'https://contoso.sharepoint.com/sites/comms',
            verbose: true,
            debug: true,
            _: []
          }
        })) {
          return Promise.resolve({
            stdout: ''
          });
        }
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/comms', title: 'Comms', type: 'CommunicationSite', debug: true } } as any, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', `Error: ${JSON.stringify(err)}`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates modern team site if existing modern team site found (no type specified)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-22T18:39:51.06",
            "CurrentChangeToken": {
              "StringValue": "1;2;113ba5b6-c737-4a6b-b1c7-2a367290057e;637470248884630000;125942029"
            },
            "CustomMasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "Description": "Team 2",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "FooterEmphasis": 0,
            "FooterEnabled": false,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "113ba5b6-c737-4a6b-b1c7-2a367290057e",
            "IsHomepageModernized": false,
            "IsMultilingual": true,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-22T18:44:16Z",
            "LastItemUserModifiedDate": "2021-01-22T18:39:57Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": false,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/team1"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/team1",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "Team 2 updated",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/team1",
            "WebTemplate": "GROUP",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      if (command === spoSiteSetCommand) {
        if (JSON.stringify(args) === JSON.stringify({
          options: {
            title: 'Team 1',
            url: 'https://contoso.sharepoint.com/sites/team1',
            verbose: false,
            debug: false,
            _: []
          }
        })) {
          return Promise.resolve({
            stdout: ''
          });
        }
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1' } } as any, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates modern team site if existing modern team site found (type specified)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-22T18:39:51.06",
            "CurrentChangeToken": {
              "StringValue": "1;2;113ba5b6-c737-4a6b-b1c7-2a367290057e;637470248884630000;125942029"
            },
            "CustomMasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "Description": "Team 2",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "FooterEmphasis": 0,
            "FooterEnabled": false,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "113ba5b6-c737-4a6b-b1c7-2a367290057e",
            "IsHomepageModernized": false,
            "IsMultilingual": true,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-22T18:44:16Z",
            "LastItemUserModifiedDate": "2021-01-22T18:39:57Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": false,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/team1"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/team1",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "Team 2 updated",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/team1",
            "WebTemplate": "GROUP",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      if (command === spoSiteSetCommand) {
        if (JSON.stringify(args) === JSON.stringify({
          options: {
            title: 'Team 1',
            url: 'https://contoso.sharepoint.com/sites/team1',
            verbose: false,
            debug: false,
            _: []
          }
        })) {
          return Promise.resolve({
            stdout: ''
          });
        }
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1', type: 'TeamSite' } } as any, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates modern communication site if existing modern communication site found (no type specified; debug)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-22T18:46:59.137",
            "CurrentChangeToken": {
              "StringValue": "1;2;7a2121e0-bf85-49dd-85c2-508e0b51b643;637471096415770000;126658766"
            },
            "CustomMasterUrl": "/sites/commsite1/_catalogs/masterpage/seattle.master",
            "Description": "",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "FooterEmphasis": 0,
            "FooterEnabled": true,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "7a2121e0-bf85-49dd-85c2-508e0b51b643",
            "IsHomepageModernized": false,
            "IsMultilingual": true,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-22T18:49:14Z",
            "LastItemUserModifiedDate": "2021-01-22T18:47:03Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/commsite1/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": true,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/commsite1"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/commsite1",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "CommSite1",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/commsite1",
            "WebTemplate": "SITEPAGEPUBLISHING",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      if (command === spoSiteSetCommand) {
        if (JSON.stringify(args) === JSON.stringify({
          options: {
            title: 'CommSite1',
            url: 'https://contoso.sharepoint.com/sites/commsite1',
            verbose: true,
            debug: true,
            _: []
          }
        })) {
          return Promise.resolve({
            stdout: ''
          });
        }
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/commsite1', title: 'CommSite1', debug: true } } as any, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates modern communication site if existing modern communication site found (type specified)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-22T18:46:59.137",
            "CurrentChangeToken": {
              "StringValue": "1;2;7a2121e0-bf85-49dd-85c2-508e0b51b643;637471096415770000;126658766"
            },
            "CustomMasterUrl": "/sites/commsite1/_catalogs/masterpage/seattle.master",
            "Description": "",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "FooterEmphasis": 0,
            "FooterEnabled": true,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "7a2121e0-bf85-49dd-85c2-508e0b51b643",
            "IsHomepageModernized": false,
            "IsMultilingual": true,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-22T18:49:14Z",
            "LastItemUserModifiedDate": "2021-01-22T18:47:03Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/commsite1/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": true,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/commsite1"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/commsite1",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "CommSite1",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/commsite1",
            "WebTemplate": "SITEPAGEPUBLISHING",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      if (command === spoSiteSetCommand) {
        if (JSON.stringify(args) === JSON.stringify({
          options: {
            title: 'CommSite1',
            url: 'https://contoso.sharepoint.com/sites/commsite1',
            verbose: false,
            debug: false,
            _: []
          }
        })) {
          return Promise.resolve({
            stdout: ''
          });
        }
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/commsite1', title: 'CommSite1', type: 'CommunicationSite' } } as any, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates classic site if an existing classic site found (type specified)', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-24T18:36:10.457",
            "CurrentChangeToken": {
              "StringValue": "1;2;b812f98e-62ff-4b0f-9cb3-91e8fe7c87b7;637471103230870000;126667243"
            },
            "CustomMasterUrl": "/sites/classic/_catalogs/masterpage/seattle.master",
            "Description": "",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": true,
            "FooterEmphasis": 0,
            "FooterEnabled": false,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "b812f98e-62ff-4b0f-9cb3-91e8fe7c87b7",
            "IsHomepageModernized": false,
            "IsMultilingual": false,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-24T18:37:34Z",
            "LastItemUserModifiedDate": "2021-01-24T18:37:21Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/classic/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": false,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/classic"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/classic",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "Classic",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/classic",
            "WebTemplate": "STS",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      if (command === spoSiteSetCommand) {
        if (JSON.stringify(args) === JSON.stringify({
          options: {
            title: 'Classic',
            url: 'https://contoso.sharepoint.com/sites/classic',
            verbose: false,
            debug: false,
            _: []
          }
        })) {
          return Promise.resolve({
            stdout: ''
          });
        }
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/classic', title: 'Classic', type: 'ClassicSite' } } as any, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`updates site's visibility and sharing options`, (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-22T18:39:51.06",
            "CurrentChangeToken": {
              "StringValue": "1;2;113ba5b6-c737-4a6b-b1c7-2a367290057e;637470248884630000;125942029"
            },
            "CustomMasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "Description": "Team 2",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "FooterEmphasis": 0,
            "FooterEnabled": false,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "113ba5b6-c737-4a6b-b1c7-2a367290057e",
            "IsHomepageModernized": false,
            "IsMultilingual": true,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-22T18:44:16Z",
            "LastItemUserModifiedDate": "2021-01-22T18:39:57Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": false,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/team1"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/team1",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "Team 2 updated",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/team1",
            "WebTemplate": "GROUP",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      if (command === spoSiteSetCommand) {
        if (JSON.stringify(args) === JSON.stringify({
          options: {
            isPublic: 'true',
            shareByEmailEnabled: 'true',
            title: 'Team 1',
            url: 'https://contoso.sharepoint.com/sites/team1',
            verbose: false,
            debug: false,
            _: []
          }
        })) {
          return Promise.resolve({
            stdout: ''
          });
        }
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1', isPublic: true, shareByEmailEnabled: true } } as any, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', `Error: ${JSON.stringify(err)}`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when validation of options for creating site failed', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.reject({
          error: new CommandError('404 FILE NOT FOUND')
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1' } } as any, (err: any) => {
      try {
        assert.notStrictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when an error has occurred when checking if a site exists at the specified URL', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.reject({
          error: new CommandError('An error has occurred')
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1' } } as any, (err: any) => {
      try {
        assert.notStrictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when the specified site type is invalid', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-24T18:36:10.457",
            "CurrentChangeToken": {
              "StringValue": "1;2;b812f98e-62ff-4b0f-9cb3-91e8fe7c87b7;637471103230870000;126667243"
            },
            "CustomMasterUrl": "/sites/classic/_catalogs/masterpage/seattle.master",
            "Description": "",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": true,
            "FooterEmphasis": 0,
            "FooterEnabled": false,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "b812f98e-62ff-4b0f-9cb3-91e8fe7c87b7",
            "IsHomepageModernized": false,
            "IsMultilingual": false,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-24T18:37:34Z",
            "LastItemUserModifiedDate": "2021-01-24T18:37:21Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/classic/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": false,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/classic"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/classic",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "Classic",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/classic",
            "WebTemplate": "STS",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/classic', title: 'Classic', type: 'Invalid' } } as any, (err: any) => {
      try {
        assert.notStrictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when a communication site expected but a team site found', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-22T18:39:51.06",
            "CurrentChangeToken": {
              "StringValue": "1;2;113ba5b6-c737-4a6b-b1c7-2a367290057e;637470248884630000;125942029"
            },
            "CustomMasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "Description": "Team 2",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "FooterEmphasis": 0,
            "FooterEnabled": false,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "113ba5b6-c737-4a6b-b1c7-2a367290057e",
            "IsHomepageModernized": false,
            "IsMultilingual": true,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-22T18:44:16Z",
            "LastItemUserModifiedDate": "2021-01-22T18:39:57Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": false,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/team1"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/team1",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "Team 2 updated",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/team1",
            "WebTemplate": "GROUP",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1', type: 'CommunicationSite' } } as any, (err: any) => {
      try {
        assert.notStrictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when no properties to update specified', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === spoWebGetCommand) {
        return Promise.resolve({
          stdout: JSON.stringify({
            "AllowRssFeeds": true,
            "AlternateCssUrl": "",
            "AppInstanceId": "00000000-0000-0000-0000-000000000000",
            "ClassicWelcomePage": null,
            "Configuration": 0,
            "Created": "2021-01-22T18:39:51.06",
            "CurrentChangeToken": {
              "StringValue": "1;2;113ba5b6-c737-4a6b-b1c7-2a367290057e;637470248884630000;125942029"
            },
            "CustomMasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "Description": "Team 2",
            "DesignPackageId": "00000000-0000-0000-0000-000000000000",
            "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
            "EnableMinimalDownload": false,
            "FooterEmphasis": 0,
            "FooterEnabled": false,
            "FooterLayout": 0,
            "HeaderEmphasis": 0,
            "HeaderLayout": 0,
            "HideTitleInHeader": false,
            "HorizontalQuickLaunch": false,
            "Id": "113ba5b6-c737-4a6b-b1c7-2a367290057e",
            "IsHomepageModernized": false,
            "IsMultilingual": true,
            "IsRevertHomepageLinkHidden": false,
            "Language": 1033,
            "LastItemModifiedDate": "2021-01-22T18:44:16Z",
            "LastItemUserModifiedDate": "2021-01-22T18:39:57Z",
            "LogoAlignment": 0,
            "MasterUrl": "/sites/team1/_catalogs/masterpage/seattle.master",
            "MegaMenuEnabled": false,
            "NavAudienceTargetingEnabled": false,
            "NoCrawl": false,
            "ObjectCacheEnabled": false,
            "OverwriteTranslationsOnChange": false,
            "ResourcePath": {
              "DecodedUrl": "https://contoso.sharepoint.com/sites/team1"
            },
            "QuickLaunchEnabled": true,
            "RecycleBinEnabled": true,
            "SearchScope": 0,
            "ServerRelativeUrl": "/sites/team1",
            "SiteLogoUrl": null,
            "SyndicationEnabled": true,
            "TenantAdminMembersCanShare": 0,
            "Title": "Team 2 updated",
            "TreeViewEnabled": false,
            "UIVersion": 15,
            "UIVersionConfigurationEnabled": false,
            "Url": "https://contoso.sharepoint.com/sites/team1",
            "WebTemplate": "GROUP",
            "WelcomePage": "SitePages/Home.aspx"
          })
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1' } } as any, (err: any) => {
      try {
        assert.notStrictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('fails validation if the specified url is a single word', async () => {
    const options: any = { url: 'site', title: 'Site' };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(typeof actual, 'string');
  });

  it('fails validation if the specified url is a server-relative URL', async () => {
    const options: any = { url: '/sites/site', title: 'Site' };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(typeof actual, 'string');
  });

  it('passes validation when all options are specified and valid', async () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/site', title: 'Site' };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, true);
  });
});