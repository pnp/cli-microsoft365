import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./orgassetslibrary-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.ORGASSETSLIBRARY_LIST, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
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
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      (command as any).getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ORGASSETSLIBRARY_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('returns a result with a thumbnail', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          { "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "8992299e-a003-4000-7686-fda36e26a53c" }, 4, { "IsNull": false }, 5, { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "fac1fa9e-e0cc-1000-077b-61deac0da407|908bed80-a04a-4433-b4a0-883d9847d110:a1214787-77d5-4b72-a96d-1c278f72bbb0nTenant", "AllowCommentsTextOnEmailEnabled": true, "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [], "AllowEditing": true, "AllowGuestUserShareToUsersNotInSiteCollection": false, "AllowLimitedAccessOnUnmanagedDevices": false, "AllowSelectSGsInODBListInTenant": null, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnFilesDisabled": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "ConditionalAccessPolicyErrorHelpLink": "", "ContentTypeSyncSiteTemplatesList": [], "CustomizedExternalSharingServiceUrl": "", "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationEnabled": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableAIPIntegration": false, "EnableAzureADB2BIntegration": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "EnablePromotedFileHandlers": true, "ExcludedFileExtensionsForSyncClient": [""], "ExternalServicesEnabled": true, "ExternalUserExpirationRequired": false, "ExternalUserExpireInDays": 60, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "GuestSharingGroupAllowListInTenant": "", "GuestSharingGroupAllowListInTenantByPrincipalIdentity": null, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": true, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "MarkNewFilesSensitiveByDefault": 0, "MobileFriendlyUrlEnabledInTenant": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "ODBSharingCapability": 2, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrgNewsSiteUrl": null, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": true, "PublicCdnOrigins": [], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 5300, "ResourceQuotaAllocated": 300, "RootSiteUrl": "https:u002fu002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1304576, "StorageQuotaAllocated": 131072000, "SyncAadB2BManagementPolicy": false, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true, "WhoCanShareAllowListInTenant": "", "WhoCanShareAllowListInTenantByPrincipalIdentity": null }, 6, {
            "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssets",
            "CentralAssetRepositoryLibraries": null,
            "OrgAssetsLibraries": {
              "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibraryCollection",
              "_Child_Items_": [{
                "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibrary",
                "DisplayName": "Site Assets",
                "FileType": "jpg",
                "LibraryUrl": {
                  "_ObjectType_": "SP.ResourcePath",
                  "DecodedUrl": "sites\u002fsitedesigns\u002fSiteAssets"
                },
                "ListId": "\/Guid(96c2e234-c996-4877-b3a6-8aebd8ab45b6)\/",
                "OrgAssetType": 1,
                "ThumbnailUrl": {
                  "_ObjectType_": "SP.ResourcePath",
                  "DecodedUrl": "SiteAssets\u002f__siteIcon__.jpg"
                },
                "UniqueId": "\/Guid(0d3c9e72-60f5-40f8-9e29-b91036f5630e)\/"
              }]
            },
            "SiteId": "\/Guid(9f0e0a96-14ec-4d4f-9b04-a8698367cd36)\/",
            "Url": {
              "_ObjectType_": "SP.ResourcePath",
              "DecodedUrl": "\u002fsites\u002fsitedesigns"
            },
            "WebId": "\/Guid(030c8d27-1bb4-4042-a252-dce8ac1e9f00)\/"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Url: '/sites/sitedesigns',
          Libraries:
            [{ DisplayName: 'Site Assets', LibraryUrl: 'sites/sitedesigns/SiteAssets', ListId: '/Guid(96c2e234-c996-4877-b3a6-8aebd8ab45b6)/', ThumbnailUrl: 'SiteAssets/__siteIcon__.jpg' }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns multiple results', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          { "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "8992299e-a003-4000-7686-fda36e26a53c" }, 4, { "IsNull": false }, 5, { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "fac1fa9e-e0cc-1000-077b-61deac0da407|908bed80-a04a-4433-b4a0-883d9847d110:a1214787-77d5-4b72-a96d-1c278f72bbb0nTenant", "AllowCommentsTextOnEmailEnabled": true, "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [], "AllowEditing": true, "AllowGuestUserShareToUsersNotInSiteCollection": false, "AllowLimitedAccessOnUnmanagedDevices": false, "AllowSelectSGsInODBListInTenant": null, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnFilesDisabled": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "ConditionalAccessPolicyErrorHelpLink": "", "ContentTypeSyncSiteTemplatesList": [], "CustomizedExternalSharingServiceUrl": "", "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationEnabled": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableAIPIntegration": false, "EnableAzureADB2BIntegration": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "EnablePromotedFileHandlers": true, "ExcludedFileExtensionsForSyncClient": [""], "ExternalServicesEnabled": true, "ExternalUserExpirationRequired": false, "ExternalUserExpireInDays": 60, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "GuestSharingGroupAllowListInTenant": "", "GuestSharingGroupAllowListInTenantByPrincipalIdentity": null, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": true, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "MarkNewFilesSensitiveByDefault": 0, "MobileFriendlyUrlEnabledInTenant": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "ODBSharingCapability": 2, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrgNewsSiteUrl": null, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": true, "PublicCdnOrigins": [], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 5300, "ResourceQuotaAllocated": 300, "RootSiteUrl": "https:u002fu002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1304576, "StorageQuotaAllocated": 131072000, "SyncAadB2BManagementPolicy": false, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true, "WhoCanShareAllowListInTenant": "", "WhoCanShareAllowListInTenantByPrincipalIdentity": null }, 6, {
            "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssets",
            "CentralAssetRepositoryLibraries": null,
            "OrgAssetsLibraries": {
              "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibraryCollection",
              "_Child_Items_": [{
                "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibrary",
                "DisplayName": "Site Assets",
                "FileType": "jpg",
                "LibraryUrl": {
                  "_ObjectType_": "SP.ResourcePath",
                  "DecodedUrl": "sites\u002fsitedesigns\u002fSiteAssets"
                },
                "ListId": "\/Guid(96c2e234-c996-4877-b3a6-8aebd8ab45b6)\/",
                "OrgAssetType": 1,
                "ThumbnailUrl": {
                  "_ObjectType_": "SP.ResourcePath",
                  "DecodedUrl": "SiteAssets\u002f__siteIcon__.jpg"
                },
                "UniqueId": "\/Guid(0d3c9e72-60f5-40f8-9e29-b91036f5630e)\/"
              }, {
                "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibrary",
                "DisplayName": "Site Assets 2",
                "FileType": "jpg",
                "LibraryUrl": {
                  "_ObjectType_": "SP.ResourcePath",
                  "DecodedUrl": "sites\u002fsitedesigns\u002fSiteAssets2"
                },
                "ListId": "\/Guid(86c2e234-c996-4877-b3a6-8aebd8ab45b6)\/",
                "OrgAssetType": 1,
                "ThumbnailUrl": null,
                "UniqueId": "\/Guid(1d3c9e72-60f5-40f8-9e29-b91036f5630e)\/"
              }]
            },
            "SiteId": "\/Guid(9f0e0a96-14ec-4d4f-9b04-a8698367cd36)\/",
            "Url": {
              "_ObjectType_": "SP.ResourcePath",
              "DecodedUrl": "\u002fsites\u002fsitedesigns"
            },
            "WebId": "\/Guid(030c8d27-1bb4-4042-a252-dce8ac1e9f00)\/"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Url: '/sites/sitedesigns', Libraries:
            [{ DisplayName: 'Site Assets', LibraryUrl: 'sites/sitedesigns/SiteAssets', ListId: '/Guid(96c2e234-c996-4877-b3a6-8aebd8ab45b6)/', ThumbnailUrl: 'SiteAssets/__siteIcon__.jpg' }, { DisplayName: 'Site Assets 2', LibraryUrl: 'sites/sitedesigns/SiteAssets2', ListId: '/Guid(86c2e234-c996-4877-b3a6-8aebd8ab45b6)/', ThumbnailUrl: null }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns a result without a thumbnail', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          { "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "8992299e-a003-4000-7686-fda36e26a53c" }, 4, { "IsNull": false }, 5, { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "fac1fa9e-e0cc-1000-077b-61deac0da407|908bed80-a04a-4433-b4a0-883d9847d110:a1214787-77d5-4b72-a96d-1c278f72bbb0nTenant", "AllowCommentsTextOnEmailEnabled": true, "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [], "AllowEditing": true, "AllowGuestUserShareToUsersNotInSiteCollection": false, "AllowLimitedAccessOnUnmanagedDevices": false, "AllowSelectSGsInODBListInTenant": null, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnFilesDisabled": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "ConditionalAccessPolicyErrorHelpLink": "", "ContentTypeSyncSiteTemplatesList": [], "CustomizedExternalSharingServiceUrl": "", "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationEnabled": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableAIPIntegration": false, "EnableAzureADB2BIntegration": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "EnablePromotedFileHandlers": true, "ExcludedFileExtensionsForSyncClient": [""], "ExternalServicesEnabled": true, "ExternalUserExpirationRequired": false, "ExternalUserExpireInDays": 60, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "GuestSharingGroupAllowListInTenant": "", "GuestSharingGroupAllowListInTenantByPrincipalIdentity": null, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": true, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "MarkNewFilesSensitiveByDefault": 0, "MobileFriendlyUrlEnabledInTenant": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "ODBSharingCapability": 2, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrgNewsSiteUrl": null, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": true, "PublicCdnOrigins": [], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 5300, "ResourceQuotaAllocated": 300, "RootSiteUrl": "https:u002fu002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1304576, "StorageQuotaAllocated": 131072000, "SyncAadB2BManagementPolicy": false, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true, "WhoCanShareAllowListInTenant": "", "WhoCanShareAllowListInTenantByPrincipalIdentity": null }, 6, {
            "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssets",
            "CentralAssetRepositoryLibraries": null,
            "OrgAssetsLibraries": {
              "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibraryCollection",
              "_Child_Items_": [{
                "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibrary",
                "DisplayName": "Site Assets",
                "FileType": "jpg",
                "LibraryUrl": {
                  "_ObjectType_": "SP.ResourcePath",
                  "DecodedUrl": "sites\u002fsitedesigns\u002fSiteAssets"
                },
                "ListId": "\/Guid(96c2e234-c996-4877-b3a6-8aebd8ab45b6)\/",
                "OrgAssetType": 1,
                "ThumbnailUrl": null,
                "UniqueId": "\/Guid(0d3c9e72-60f5-40f8-9e29-b91036f5630e)\/"
              }]
            },
            "SiteId": "\/Guid(9f0e0a96-14ec-4d4f-9b04-a8698367cd36)\/",
            "Url": {
              "_ObjectType_": "SP.ResourcePath",
              "DecodedUrl": "\u002fsites\u002fsitedesigns"
            },
            "WebId": "\/Guid(030c8d27-1bb4-4042-a252-dce8ac1e9f00)\/"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          Url: '/sites/sitedesigns',
          Libraries:
            [{ DisplayName: 'Site Assets', LibraryUrl: 'sites/sitedesigns/SiteAssets', ListId: '/Guid(96c2e234-c996-4877-b3a6-8aebd8ab45b6)/', ThumbnailUrl: null }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns a result as json', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          { "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "8992299e-a003-4000-7686-fda36e26a53c" }, 4, { "IsNull": false }, 5, { "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "fac1fa9e-e0cc-1000-077b-61deac0da407|908bed80-a04a-4433-b4a0-883d9847d110:a1214787-77d5-4b72-a96d-1c278f72bbb0nTenant", "AllowCommentsTextOnEmailEnabled": true, "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [], "AllowEditing": true, "AllowGuestUserShareToUsersNotInSiteCollection": false, "AllowLimitedAccessOnUnmanagedDevices": false, "AllowSelectSGsInODBListInTenant": null, "ApplyAppEnforcedRestrictionsToAdHocRecipients": true, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnFilesDisabled": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "ConditionalAccessPolicyErrorHelpLink": "", "ContentTypeSyncSiteTemplatesList": [], "CustomizedExternalSharingServiceUrl": "", "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisabledWebPartIds": null, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationEnabled": false, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableAIPIntegration": false, "EnableAzureADB2BIntegration": false, "EnableGuestSignInAcceleration": false, "EnableMinimumVersionRequirement": true, "EnablePromotedFileHandlers": true, "ExcludedFileExtensionsForSyncClient": [""], "ExternalServicesEnabled": true, "ExternalUserExpirationRequired": false, "ExternalUserExpireInDays": 60, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "GuestSharingGroupAllowListInTenant": "", "GuestSharingGroupAllowListInTenantByPrincipalIdentity": null, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsHubSitesMultiGeoFlightEnabled": true, "IsMultiGeo": false, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "LimitedAccessFileType": 1, "MarkNewFilesSensitiveByDefault": 0, "MobileFriendlyUrlEnabledInTenant": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "ODBSharingCapability": 2, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrgNewsSiteUrl": null, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": true, "PublicCdnOrigins": [], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": -1, "ResourceQuota": 5300, "ResourceQuotaAllocated": 300, "RootSiteUrl": "https:u002fu002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": false, "ShowEveryoneClaim": false, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SocialBarOnSitePagesDisabled": false, "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1304576, "StorageQuotaAllocated": 131072000, "SyncAadB2BManagementPolicy": false, "SyncPrivacyProfileProperties": true, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true, "WhoCanShareAllowListInTenant": "", "WhoCanShareAllowListInTenantByPrincipalIdentity": null }, 6, {
            "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssets",
            "CentralAssetRepositoryLibraries": null,
            "OrgAssetsLibraries": {
              "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibraryCollection",
              "_Child_Items_": [{
                "_ObjectType_": "Microsoft.SharePoint.Administration.OrgAssetsLibrary",
                "DisplayName": "Site Assets",
                "FileType": "jpg",
                "LibraryUrl": {
                  "_ObjectType_": "SP.ResourcePath",
                  "DecodedUrl": "sites\u002fsitedesigns\u002fSiteAssets"
                },
                "ListId": "\/Guid(96c2e234-c996-4877-b3a6-8aebd8ab45b6)\/",
                "OrgAssetType": 1,
                "ThumbnailUrl": {
                  "_ObjectType_": "SP.ResourcePath",
                  "DecodedUrl": "SiteAssets\u002f__siteIcon__.jpg"
                },
                "UniqueId": "\/Guid(0d3c9e72-60f5-40f8-9e29-b91036f5630e)\/"
              }]
            },
            "SiteId": "\/Guid(9f0e0a96-14ec-4d4f-9b04-a8698367cd36)\/",
            "Url": {
              "_ObjectType_": "SP.ResourcePath",
              "DecodedUrl": "\u002fsites\u002fsitedesigns"
            },
            "WebId": "\/Guid(030c8d27-1bb4-4042-a252-dce8ac1e9f00)\/"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, verbose: true, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(JSON.stringify({
          Url: '/sites/sitedesigns',
          Libraries:
            [{ DisplayName: 'Site Assets', LibraryUrl: 'sites/sitedesigns/SiteAssets', ListId: '/Guid(96c2e234-c996-4877-b3a6-8aebd8ab45b6)/', ThumbnailUrl: 'SiteAssets/__siteIcon__.jpg' }]
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles no library set correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{
          "SchemaVersion": "15.0.0.0",
          "LibraryVersion": "16.0.19131.12010",
          "ErrorInfo": null,
          "TraceCorrelationId": "46b3fa9e-704c-1000-1fc5-a24124d1d3f0"
        }, 4, {
          "IsNull": false
        }, 5, {
          "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant",
          "_ObjectIdentity_": "46b3fa9e-704c-1000-1fc5-a24124d1d3f0|908bed80-a04a-4433-b4a0-883d9847d110:a1214787-77d5-4b72-a96d-1c278f72bbb0nTenant",
          "AllowCommentsTextOnEmailEnabled": true,
          "AllowDownloadingNonWebViewableFiles": true,
          "AllowedDomainListForSyncClient": [

          ],
          "AllowEditing": true,
          "AllowGuestUserShareToUsersNotInSiteCollection": false,
          "AllowLimitedAccessOnUnmanagedDevices": false,
          "AllowSelectSGsInODBListInTenant": null,
          "ApplyAppEnforcedRestrictionsToAdHocRecipients": true,
          "BccExternalSharingInvitations": false,
          "BccExternalSharingInvitationsList": null,
          "BlockAccessOnUnmanagedDevices": false,
          "BlockDownloadOfAllFilesForGuests": false,
          "BlockDownloadOfAllFilesOnUnmanagedDevices": false,
          "BlockDownloadOfViewableFilesForGuests": false,
          "BlockDownloadOfViewableFilesOnUnmanagedDevices": false,
          "BlockMacSync": false,
          "CommentsOnFilesDisabled": false,
          "CommentsOnSitePagesDisabled": false,
          "CompatibilityRange": "15,15",
          "ConditionalAccessPolicy": 0,
          "ConditionalAccessPolicyErrorHelpLink": "",
          "ContentTypeSyncSiteTemplatesList": [

          ],
          "CustomizedExternalSharingServiceUrl": "",
          "DefaultLinkPermission": 0,
          "DefaultSharingLinkType": 3,
          "DisabledWebPartIds": null,
          "DisableReportProblemDialog": false,
          "DisallowInfectedFileDownload": false,
          "DisplayNamesOfFileViewers": true,
          "DisplayStartASiteOption": true,
          "EmailAttestationEnabled": false,
          "EmailAttestationReAuthDays": 30,
          "EmailAttestationRequired": false,
          "EnableAIPIntegration": false,
          "EnableAzureADB2BIntegration": false,
          "EnableGuestSignInAcceleration": false,
          "EnableMinimumVersionRequirement": true,
          "EnablePromotedFileHandlers": true,
          "ExcludedFileExtensionsForSyncClient": [
            ""
          ],
          "ExternalServicesEnabled": true,
          "ExternalUserExpirationRequired": false,
          "ExternalUserExpireInDays": 60,
          "FileAnonymousLinkType": 2,
          "FilePickerExternalImageSearchEnabled": true,
          "FolderAnonymousLinkType": 2,
          "GuestSharingGroupAllowListInTenant": "",
          "GuestSharingGroupAllowListInTenantByPrincipalIdentity": null,
          "HideSyncButtonOnODB": false,
          "IPAddressAllowList": "",
          "IPAddressEnforcement": false,
          "IPAddressWACTokenLifetime": 15,
          "IsHubSitesMultiGeoFlightEnabled": true,
          "IsMultiGeo": false,
          "IsUnmanagedSyncClientForTenantRestricted": false,
          "IsUnmanagedSyncClientRestrictionFlightEnabled": true,
          "LegacyAuthProtocolsEnabled": true,
          "LimitedAccessFileType": 1,
          "MarkNewFilesSensitiveByDefault": 0,
          "MobileFriendlyUrlEnabledInTenant": true,
          "NoAccessRedirectUrl": null,
          "NotificationsInOneDriveForBusinessEnabled": true,
          "NotificationsInSharePointEnabled": true,
          "NotifyOwnersWhenInvitationsAccepted": true,
          "NotifyOwnersWhenItemsReshared": true,
          "ODBAccessRequests": 0,
          "ODBMembersCanShare": 0,
          "ODBSharingCapability": 2,
          "OfficeClientADALDisabled": false,
          "OneDriveForGuestsEnabled": false,
          "OneDriveStorageQuota": 1048576,
          "OptOutOfGrooveBlock": false,
          "OptOutOfGrooveSoftBlock": false,
          "OrgNewsSiteUrl": null,
          "OrphanedPersonalSitesRetentionPeriod": 30,
          "OwnerAnonymousNotification": true,
          "PermissiveBrowserFileHandlingOverride": false,
          "PreventExternalUsersFromResharing": false,
          "ProvisionSharedWithEveryoneFolder": false,
          "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF",
          "PublicCdnEnabled": true,
          "PublicCdnOrigins": [

          ],
          "RequireAcceptingAccountMatchInvitedAccount": false,
          "RequireAnonymousLinksExpireInDays": -1,
          "ResourceQuota": 5300,
          "ResourceQuotaAllocated": 300,
          "RootSiteUrl": "https:u002fu002fcontoso.sharepoint.com",
          "SearchResolveExactEmailOrUPN": false,
          "SharingAllowedDomainList": null,
          "SharingBlockedDomainList": null,
          "SharingCapability": 2,
          "SharingDomainRestrictionMode": 0,
          "ShowAllUsersClaim": false,
          "ShowEveryoneClaim": false,
          "ShowEveryoneExceptExternalUsersClaim": true,
          "ShowNGSCDialogForSyncOnODB": true,
          "ShowPeoplePickerSuggestionsForGuestUsers": false,
          "SignInAccelerationDomain": "",
          "SocialBarOnSitePagesDisabled": false,
          "SpecialCharactersStateInFileFolderNames": 1,
          "StartASiteFormUrl": null,
          "StorageQuota": 1304576,
          "StorageQuotaAllocated": 131072000,
          "SyncAadB2BManagementPolicy": false,
          "SyncPrivacyProfileProperties": true,
          "UseFindPeopleInPeoplePicker": false,
          "UsePersistentCookiesForExplorerView": false,
          "UserVoiceForFeedbackEnabled": true,
          "WhoCanShareAllowListInTenant": "",
          "WhoCanShareAllowListInTenantByPrincipalIdentity": null
        }, 6, null]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true
      }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith('No libraries in Organization Assets'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error getting request', (done) => {
    const svcListRequest = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
            }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true
      }
    }, (err?: any) => {
      try {
        assert(svcListRequest.called);
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});
