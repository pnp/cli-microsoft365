import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./site-remove');

describe(commands.SITE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    let futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: futureDate.toISOString() }); });
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
    requests = [];
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
  });

  afterEach(() => {
    (command as any).currentContext = undefined;
    Utils.restore([
      request.post,
      request.delete,
      global.setTimeout,
      (command as any).ensureFormDigest,
      Cli.prompt
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
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', debug: true, verbose: true } }, () => {
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', debug: true, verbose: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
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
    const actual = command.validate({
      options: {
        url: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url is not a valid SharePoint url', () => {
    const actual = command.validate({
      options: {
        url: 'http://contoso'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the required options are correct', () => {
    const actual = command.validate({
      options: {
        url: 'https://contoso.sharepoint.com/sites/demosite'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('removes site. doesn\'t wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true } }, () => {
      try {
        assert(loggerSpy.notCalled);
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: true, skipRecycleBin: true } } as any, (error: any) => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, skipRecycleBin: true } } as any, (error: any) => {
      try {
        assert(loggerSpy.notCalled);
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', fromRecycleBin: true, confirm: true, debug: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', fromRecycleBin: true, confirm: true } }, () => {
      try {
        assert(loggerSpy.notCalled);
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
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
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', fromRecycleBin: true, confirm: true, debug: true, wait: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', fromRecycleBin: true, confirm: true, debug: true, wait: true } } as any, (err?: any) => {
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: true, wait: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, verbose: true, wait: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: false, wait: true } }, () => {
      try {
        assert(loggerSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the site - Groupified Site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demositeGrouped</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(58587cc9-560c-4adb-a849-e669bd37c5f8)/"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/58587cc9-560c-4adb-a849-e669bd37c5f8') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demositeGrouped', debug: true, verbose: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
        Utils.restore(request.delete);
      }
    });
  });

  it('removes the site - Groupified Site - Entered with \'fromRecycleBin\' OR \'skipRecycleBin\' ', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demositeGrouped</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(58587cc9-560c-4adb-a849-e669bd37c5f8)/"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/58587cc9-560c-4adb-a849-e669bd37c5f8') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demositeGrouped', debug: true, verbose: true, skipRecycleBin: true } }, () => {
      try {
        assert(loggerSpy.calledWith(chalk.green('DONE')));
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

  it('removes site. wait for completion, error occured', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
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

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: false, wait: true } } as any, (err?: any) => {
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
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": null,
              "TraceCorrelationId": "10f1829f-d000-0000-5962-1110d33e2cf2"
            },
            4,
            {
              "IsNull": false
            },
            5,
            {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
              "_ObjectIdentity_": "10f1829f-d000-0000-5962-1110d33e2cf2|908bed80-a04a-4433-b4a0-883d9847d110:095efa67-57fa-40c7-b7cc-e96dc3e5780c\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fdemosite",
              "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/"
            }
          ]));
        }
      }
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demosite</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demosite', confirm: true, debug: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Get Group ID - Error Occurred', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.data === `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/demositeinvalid</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.20530.12001",
              "ErrorInfo": {
                "ErrorMessage": "Cannot get site https://contoso.sharepoint.com/sites/demositeinvalid.",
                "ErrorValue": null,
                "TraceCorrelationId": "3929839f-9018-0000-5518-a12b0af612a8",
                "ErrorCode": -1,
                "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              },
              "TraceCorrelationId": "3929839f-9018-0000-5518-a12b0af612a8"
            }
          ]));
        }
      }

      return Promise.reject('Cannot get site https://contoso.sharepoint.com/sites/demositeinvalid.');
    });
    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/demositeinvalid', confirm: true, debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Cannot get site https://contoso.sharepoint.com/sites/demositeinvalid.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });
});