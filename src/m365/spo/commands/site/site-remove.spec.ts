import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./site-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import config from '../../../../config';
import * as chalk from 'chalk';

describe(commands.SITE_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdClassicInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    let futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: futureDate.toISOString() }); }); 
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
    cmdClassicInstance = {
      commandWrapper: {
        command: commands.SITE_REMOVE
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    requests = [];
  });

  afterEach(() => {
    (command as any).currentContext = undefined;
    Utils.restore([
      request.post,
      global.setTimeout,
      (command as any).ensureFormDigest
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
    assert.strictEqual(command.name.startsWith(commands.SITE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('aborts removing site when prompt not confirmed', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', debug: true, verbose: true } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('runs with alias and aborts when prompt not confirmed', (done) => {
    cmdClassicInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdClassicInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', debug: true, verbose: true } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the site when prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', debug: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('fails validation if the url is not a valid url', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url is not a valid SharePoint url', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'http://contoso'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the required options are correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/demosite'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('removes site. doesn\'t wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site, skip recyclebin doesn\'t wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveDeletedSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: true, skipRecycleBin: true } }, (error: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site, skip recyclebin doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveDeletedSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, skipRecycleBin: true } }, (error: any) => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site from recyclebin doesn\'t wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveDeletedSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', fromRecycleBin: true, confirm: true, debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site from recyclebin doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8015.1210", "ErrorInfo": null, "TraceCorrelationId": "5eda879e-90d5-6000-d611-e6bfd5acde9f"
            }, 12, {
              "IsNull": false
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 2, "DefaultSharingLinkType": 2, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 1, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 1, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": false, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 5242880, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 6300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1355776, "StorageQuotaAllocated": 135266304, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "5eda879e-90d5-6000-d611-e6bfd5acde9f|908bed80-a04a-4433-b4a0-883d9847d110:2ca3eaa5-140f-4175-9563-1172edf9f339\nSpoOperation\nRemoveDeletedSite\n636707032254311675\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', fromRecycleBin: true, confirm: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site from recycle bin, wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096914"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096914|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));

        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', fromRecycleBin: true, confirm: true, debug: true, wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site from recyclebin, wait for completion, error occured', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', fromRecycleBin: true, confirm: true, debug: true, wait: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site. wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096914"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096914|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: true, wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site. wait for completion (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096914"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096914|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, verbose: true, wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site. wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096914"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096914|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));

        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: false, wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site. wait for completion, error occured', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));

        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: false, wait: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site, error occured', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
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
});