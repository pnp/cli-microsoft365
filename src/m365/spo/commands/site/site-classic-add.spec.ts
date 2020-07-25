import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import config from '../../../../config';
import auth from '../../../../Auth';
const command: Command = require('./site-classic-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.SITE_CLASSIC_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.SITE_CLASSIC_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates classic site with minimal options. doesn\'t wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with full options. doesn\'t wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">300</Property><Property Name="StorageWarningLevel" Type="Int64">275</Property><Property Name="Template" Type="String">PUBLISHING#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">100</Property><Property Name="UserCodeWarningLevel" Type="Double">90</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', lcid: 1033, webTemplate: 'PUBLISHING#0', resourceQuota: 100, resourceQuotaWarningLevel: 90, storageQuota: 300, storageQuotaWarningLevel: 275 } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fsite4.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "c340489e-80cc-5000-c5b4-01b2ce71e9bf"
            }, 197, {
              "IsNull": false
            }, 199, {
              "IsNull": false
            }, 200, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "c340489e-80cc-5000-c5b4-01b2ce71e9bf|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "AllowDownloadingNonWebViewableFiles": true, "AllowEditing": true, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 1, "DisableAppViews": 2, "DisableCompanyWideSharingLinks": 2, "DisableFlows": 2, "HasHolds": false, "LastContentModifiedDate": "\/Date(2018,1,7,19,9,58,513)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "admin@contoso.onmicrosoft.com", "OwnerEmail": "admin@contoso.onmicrosoft.com", "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 2, "SharingAllowedDomainList": "", "SharingBlockedDomainList": "", "SharingCapability": 0, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 2, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 4, "Title": "Team", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0, "UserCodeWarningLevel": 0, "WebsCount": 1
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site is being deleted', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "c340489e-80cc-5000-c5b4-01b2ce71e9bf"
            }, 197, {
              "IsNull": false
            }, 199, {
              "IsNull": false
            }, 200, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "c340489e-80cc-5000-c5b4-01b2ce71e9bf|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "AllowDownloadingNonWebViewableFiles": true, "AllowEditing": true, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 1, "DisableAppViews": 2, "DisableCompanyWideSharingLinks": 2, "DisableFlows": 2, "HasHolds": false, "LastContentModifiedDate": "\/Date(2018,1,7,19,9,58,513)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "admin@contoso.onmicrosoft.com", "OwnerEmail": "admin@contoso.onmicrosoft.com", "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 2, "SharingAllowedDomainList": "", "SharingBlockedDomainList": "", "SharingCapability": 0, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Recycled", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 2, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 4, "Title": "Team", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0, "UserCodeWarningLevel": 0, "WebsCount": 1
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, error while checking if site exists in the recycle bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, error while checking if deleted site exists in the recycle bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fsite4.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }
      }

      if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
              "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Exception"
            }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site exists in the recycle bin', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site exists in the recycle bin but is invalid', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Invalid", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site exists in the recycle bin (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site exists in the recycle bin. error while deleting the site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site doesn\'t exist. refreshes expired token', (done) => {
    Utils.restore((command as any).ensureFormDigest);

    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fsite4.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. doesn\'t wait for completion. remove deleted site, site doesn\'t exist. refreshes expired token (debug)', (done) => {
    Utils.restore((command as any).ensureFormDigest);
    //sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: -1 }); });
    let pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(command as any, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate.toISOString() }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fsite4.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. wait for completion. operation immediately completed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. remove deleted site. site exists in the recycle bin. wait for completion. operation immediately completed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', wait: true, removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
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

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', wait: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. remove deleted site. site exists in the recycle bin. wait for completion', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;cb09f194-0ee7-4c48-a44f-8c112fff4d4e" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
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

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', wait: true, removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. remove deleted site. site exists in the recycle bin. wait for completion (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;cb09f194-0ee7-4c48-a44f-8c112fff4d4e" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
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

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', wait: true, removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. remove deleted site. site exists in the recycle bin. wait for completion (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;cb09f194-0ee7-4c48-a44f-8c112fff4d4e" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
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

    cmdInstance.action({ options: { debug: false, verbose: true, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', wait: true, removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. remove deleted site. site exists in the recycle bin. wait for completion. error while polling', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;cb09f194-0ee7-4c48-a44f-8c112fff4d4e" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
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

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', wait: true, removeDeletedSite: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates classic site with minimal options. remove deleted site. site exists in the recycle bin. wait for completion two rounds', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]));
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;RemoveDeletedSite&#xA;636536266495764941&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;cb09f194-0ee7-4c48-a44f-8c112fff4d4e" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }

        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // not done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 5000
            }
          ]));
        }

        // done
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536251347192220&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
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

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com', wait: true, removeDeletedSite: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in the request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com&gt;</Property><Property Name="StorageMaximumLevel" Type="Int64">300</Property><Property Name="StorageWarningLevel" Type="Int64">275</Property><Property Name="Template" Type="String">PUBLISHING#0&gt;</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team&gt;</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team&gt;</Property><Property Name="UserCodeMaximumLevel" Type="Double">100</Property><Property Name="UserCodeWarningLevel" Type="Double">90</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com/sites/team>', title: 'Team>', timeZone: 4, owner: 'admin@contoso.com>', lcid: 1033, webTemplate: 'PUBLISHING#0>', resourceQuota: 100, resourceQuotaWarningLevel: 90, storageQuota: 300, storageQuotaWarningLevel: 275 } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when creating site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">admin@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">Team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "A site already exists at url https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "c340489e-70f6-5000-c5b4-00bd039e3bf9", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "c340489e-70f6-5000-c5b4-00bd039e3bf9"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com/sites/team', title: 'Team', timeZone: 4, owner: 'admin@contoso.com' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("A site already exists at url https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.")));
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

  it('fails validation if the url is not a valid url', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'abc', title: 'Team', timeZone: 4, owner: 'admin@contoso.com'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url is not a valid SharePoint url', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'http://contoso', title: 'Team', timeZone: 4, owner: 'admin@contoso.com'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the timeZone is not a number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 'a'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the lcid is not a number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, lcid: 'a'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceQuota is not a number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, resourceQuota: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceQuotaWarningLevel is not a number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, resourceQuotaWarningLevel: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceQuotaWarningLevel is specified without resourceQuota', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, resourceQuotaWarningLevel: 10
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceQuotaWarningLevel is greater than resourceQuota', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, resourceQuotaWarningLevel: 10, resourceQuota: 5
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageQuota is not a number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, storageQuota: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageQuotaWarningLevel is not a number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, storageQuotaWarningLevel: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageQuotaWarningLevel is specified without storageQuota', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, storageQuotaWarningLevel: 10
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageQuotaWarningLevel is greater than storageQuota', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4, storageQuotaWarningLevel: 10, storageQuota: 5
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the required options are correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if all options are correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/team', title: 'Team',
        owner: 'admin@contoso.com', timeZone: 4,
        lcid: 1033, resourceQuota: 100, resourceQuotaWarningLevel: 90,
        storageQuota: 100, storageQuotaWarningLevel: 90
      }
    });
    assert.strictEqual(actual, true);
  });
});