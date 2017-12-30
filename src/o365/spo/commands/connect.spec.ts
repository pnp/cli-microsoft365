import commands from '../commands';
import Command, { CommandValidate, CommandCancel, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth, { Site } from '../SpoAuth';
const command: Command = require('./connect');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import config from '../../../config';
import Utils from '../../../Utils';

describe(commands.CONNECT, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(auth, 'clearSiteConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeSiteConnectionInfo').callsFake(() => Promise.resolve());
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          opts.body) {
          // tenant info request
          if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
            return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7011.1204", "ErrorInfo": null, "TraceCorrelationId": "f6ca259e-507f-4000-8546-cf9cd0cdfea5" }, 4, { "IsNull": false }, 5, { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "f6ca259e-507f-4000-8546-cf9cd0cdfea5|908bed80-a04a-4433-b4a0-883d9847d110:c7e66393-6063-455a-8b3f-9c1cc6b53033\\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [""], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": false, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 2500, "RootSiteUrl": "https://m365x755181.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 235929600, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true }]));
          }
        }
      }

      return Promise.reject('Invalid request');
    });
  });

  beforeEach(() => {
    vorpal = require('../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    sinon.stub(auth.site, 'disconnect').callsFake(() => { });
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth,
      auth.clearSiteConnectionInfo,
      auth.storeSiteConnectionInfo,
      request.post
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.CONNECT), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {}, url: 'https://contoso-admin.sharepoint.com' }, () => {
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
    cmdInstance.action({ options: {}, url: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert.equal(telemetry.name, command.name);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('connects to a tenant admin site', (done) => {
    auth.site = new Site();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false }, url: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert(auth.site.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('connects to a tenant admin site (debug)', (done) => {
    auth.site = new Site();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, url: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert(auth.site.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('connects to a regular SharePoint site', (done) => {
    auth.site = new Site();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false }, url: 'https://contoso.sharepoint.com' }, () => {
      try {
        assert(auth.site.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('connects to a regular SharePoint site (debug)', (done) => {
    auth.site = new Site();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, url: 'https://contoso.sharepoint.com' }, () => {
      try {
        assert(auth.site.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('accepts valid SharePoint Online URL', () => {
    const actual = (command.validate() as CommandValidate)({ url: 'https://contoso.sharepoint.com' });
    assert.equal(actual, true);
  });

  it('rejects invalid SharePoint Online URL', () => {
    const url = 'https://contoso.com';
    const actual = (command.validate() as CommandValidate)({ url: url });
    assert.equal(actual, `${url} is not a valid SharePoint Online site URL`);
  });

  it('can be cancelled', () => {
    assert(command.cancel());
  });

  it('clears pending connection on cancel', () => {
    auth.interval = {} as any;
    Utils.restore(global.clearInterval);
    const clearIntervalSpy = sinon.spy(global, 'clearInterval');
    (command.cancel() as CommandCancel)();
    assert(clearIntervalSpy.called);
  });

  it('doesn\'t fail when no cancelled and no connection was pending', () => {
    auth.interval = undefined as any;
    Utils.restore(global.clearInterval);
    const clearIntervalSpy = sinon.spy(global, 'clearInterval');
    (command.cancel() as CommandCancel)();
    assert.equal(clearIntervalSpy.called, false);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => {},
      prompt: () => {},
      helpInformation: () => {}
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
    assert(find.calledWith(commands.CONNECT));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => {},
      helpInformation: () => {}
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => {});
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token when connecting to a tenant-admin site', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false }, url: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles lack of valid access token when connecting to a tenant-admin site (debug)', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, url: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(auth.ensureAccessToken);
      }
    });
  });

  it('correctly handles lack of valid access token when connecting to a regular site', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false }, url: 'https://contoso.sharepoint.com' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles lack of valid access token when connecting to a regular site (debug)', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, url: 'https://contoso.sharepoint.com' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(auth.ensureAccessToken);
      }
    });
  });

  it('correctly handles error when clearing persisted auth information', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve('ABC'));
    Utils.restore(auth.clearSiteConnectionInfo);
    sinon.stub(auth, 'clearSiteConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {}, url: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          auth.clearSiteConnectionInfo,
          auth.ensureAccessToken
        ]);
      }
    });
  });

  it('correctly handles error when clearing persisted auth information (debug)', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve('ABC'));
    Utils.restore(auth.clearSiteConnectionInfo);
    sinon.stub(auth, 'clearSiteConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true }, url: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          auth.clearSiteConnectionInfo,
          auth.ensureAccessToken
        ]);
      }
    });
  });
});