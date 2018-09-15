# spo tenant settings set

Sets tenant global setting

## Usage

```sh
spo tenant settings set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging
`--minCompatibilityLevel [minCompatibilityLevel]`|Specifies the lower bound on the compatibility level for new sites'
`--maxCompatibilityLevel [maxCompatibilityLevel]`|Specifies the upper bound on the compatibility level for new sites'
`--externalServicesEnabled [externalServicesEnabled]`|Enables external services for a tenant. External services are defined as services that are not in the Office 365 datacenters. Allowed values true,false
`--noAccessRedirectUrl [noAccessRedirectUrl]`|Specifies the URL of the redirected site for those site collections which have the locked state "NoAccess"'
`--sharingCapability [sharingCapability]`|Determines what level of sharing is available for the site. The valid values are: ExternalUserAndGuestSharing (default) - External user sharing (share by email) and guest link sharing are both enabled. Disabled - External user sharing (share by email) and guest link sharing are both disabled. ExternalUserSharingOnly - External user sharing (share by email) is enabled, but guest link sharing is disabled. Allowed values Disabled,ExternalUserSharingOnly,ExternalUserAndGuestSharing,ExistingExternalUserSharingOnly
`--displayStartASiteOption [displayStartASiteOption]`|Determines whether tenant users see the Start a Site menu option. Allowed values true,false
`--startASiteFormUrl [startASiteFormUrl]`|Specifies URL of the form to load in the Start a Site dialog. The valid values are: "" (default) - Blank by default, this will also remove or clear any value that has been set. Full URL - Example:"https://contoso.sharepoint.com/path/to/form"'
`--showEveryoneClaim [showEveryoneClaim]`|Enables the administrator to hide the Everyone claim in the People Picker. When users share an item with Everyone, it is accessible to all authenticated users in the tenant\'s Azure Active Directory, including any active external users who have previously accepted invitations. Note, that some SharePoint system resources such as templates and pages are required to be shared to Everyone and this type of sharing does not expose any user data or metadata. Allowed values true,false
`--showAllUsersClaim [showAllUsersClaim]`|Enables the administrator to hide the All Users claim groups in People Picker. When users share an item with "All Users (x)", it is accessible to all organization members in the tenant\'s Azure Active Directory who have authenticated with via this method. When users share an item with "All Users (x)" it is accessible to all organtization members in the tenant that used NTLM to authentication with SharePoint. Allowed values true,false
`--showEveryoneExceptExternalUsersClaim [showEveryoneExceptExternalUsersClaim]`|Enables the administrator to hide the "Everyone except external users" claim in the People Picker. When users share an item with "Everyone except external users", it is accessible to all organization members in the tenant\'s Azure Active Directory, but not to any users who have previously accepted invitations. Allowed values true,false
`--searchResolveExactEmailOrUPN [searchResolveExactEmailOrUPN]`|Removes the search capability from People Picker. Note, recently resolved names will still appear in the list until browser cache is cleared or expired. SharePoint Administrators will still be able to use starts with or partial name matching when enabled. Allowed values true,false
`--officeClientADALDisabled [officeClientADALDisabled]`|When set to true this will disable the ability to use Modern Authentication that leverages ADAL across the tenant. Allowed values true,false
`--legacyAuthProtocolsEnabled [legacyAuthProtocolsEnabled]`|By default this value is set to true. Setting this parameter prevents Office clients using non-modern authentication protocols from accessing SharePoint Online resources. A value of true - Enables Office clients using non-modern authentication protocols(such as, Forms-Based Authentication (FBA) or Identity Client Runtime Library (IDCRL)) to access SharePoint resources. Allowed values true,false
`--requireAcceptingAccountMatchInvitedAccount [requireAcceptingAccountMatchInvitedAccount]`|Ensures that an external user can only accept an external sharing invitation with an account matching the invited email address. Administrators who desire increased control over external collaborators should consider enabling this feature. Allowed values true,false
`--provisionSharedWithEveryoneFolder [provisionSharedWithEveryoneFolder]`|Creates a Shared with Everyone folder in every user\'s new OneDrive for Business document library. The valid values are: True (default) - The Shared with Everyone folder is created. False - No folder is created when the site and OneDrive for Business document library is created. Allowed values true,false
`--signInAccelerationDomain [signInAccelerationDomain]`|Specifies the home realm discovery value to be sent to Azure Active Directory (AAD) during the user sign-in process. When the organization uses a third-party identity provider, this prevents the user from seeing the Azure Active Directory Home Realm Discovery web page and ensures the user only sees their company\'s Identity Provider\'s portal. This value can also be used with Azure Active Directory Premium to customize the Azure Active Directory login page. Acceleration will not occur on site collections that are shared externally. This value should be configured with the login domain that is used by your company (that is, example@contoso.com). If your company has multiple third-party identity providers, configuring the sign-in acceleration value will break sign-in for your organization. The valid values are: "" (default) - Blank by default, this will also remove or clear any value that has been set. Login Domain - For example: "contoso.com". No value assigned by default'
`--enableGuestSignInAcceleration [enableGuestSignInAcceleration]`|Accelerates guest-enabled site collections as well as member-only site collections when the SignInAccelerationDomain parameter is set. Allowed values true,false
`--usePersistentCookiesForExplorerView [usePersistentCookiesForExplorerView]`|Lets SharePoint issue a special cookie that will allow this feature to work even when "Keep Me Signed In" is not selected. "Open with Explorer" requires persisted cookies to operate correctly. When the user does not select "Keep Me Signed in" at the time of sign -in, "Open with Explorer" will fail. This special cookie expires after 30 minutes and cannot be cleared by closing the browser or signing out of SharePoint Online.To clear this cookie, the user must log out of their Windows session. The valid values are: False(default) - No special cookie is generated and the normal Office 365 sign -in length / timing applies. True - Generates a special cookie that will allow "Open with Explorer" to function if the "Keep Me Signed In" box is not checked at sign -in. Allowed values true,false
`--bccExternalSharingInvitations [bccExternalSharingInvitations]`|When the feature is enabled, all external sharing invitations that are sent will blind copy the e-mail messages listed in the BccExternalSharingsInvitationList. Allowed values true,false
`--bccExternalSharingInvitationsList [bccExternalSharingInvitationsList]`|Specifies a list of e-mail addresses to be BCC\'d when the BCC for External Sharing feature is enabled. Multiple addresses can be specified by creating a comma separated list with no spaces'
`--userVoiceForFeedbackEnabled [userVoiceForFeedbackEnabled]`|Enables or disables the User Voice Feedback button. Allowed values true,false
`--publicCdnEnabled [publicCdnEnabled]`|Enables or disables the publish CDN. Allowed values true,false
`--publicCdnAllowedFileTypes [publicCdnAllowedFileTypes]`|Sets public CDN allowed file types'
`--requireAnonymousLinksExpireInDays [requireAnonymousLinksExpireInDays]`|Specifies all anonymous links that have been created (or will be created) will expire after the set number of days. To remove the expiration requirement, set the value to zero (0)'
`--sharingAllowedDomainList [sharingAllowedDomainList]`|Specifies a list of email domains that is allowed for sharing with the external collaborators. Use the space character as the delimiter for entering multiple values. For example, "contoso.com fabrikam.com"'
`--sharingBlockedDomainList [sharingBlockedDomainList]`|Specifies a list of email domains that is blocked or prohibited for sharing with the external collaborators. Use space character as the delimiter for entering multiple values. For example, "contoso.com fabrikam.com"'
`--sharingDomainRestrictionMode [sharingDomainRestrictionMode]`|Specifies the external sharing mode for domains. Allowed values None,AllowList,BlockList
`--oneDriveStorageQuota [oneDriveStorageQuota]`|Sets a default OneDrive for Business storage quota for the tenant. It will be used for new OneDrive for Business sites created. A typical use will be to reduce the amount of storage associated with OneDrive for Business to a level below what the License entitles the users. For example, it could be used to set the quota to 10 gigabytes (GB) by default'
`--oneDriveForGuestsEnabled [oneDriveForGuestsEnabled]`|Lets OneDrive for Business creation for administrator managed guest users. Administrator managed Guest users use credentials in the resource tenant to access the resources. Allowed values true,false
`--iPAddressEnforcement [iPAddressEnforcement]`|Allows access from network locations that are defined by an administrator. The values are true and false. The default value is false which means the setting is disabled. Before the iPAddressEnforcement parameter is set, make sure you add a valid IPv4 or IPv6 address to the iPAddressAllowList parameter. Allowed values true,false
`--iPAddressAllowList [iPAddressAllowList]`|Configures multiple IP addresses or IP address ranges (IPv4 or IPv6). Use commas to separate multiple IP addresses or IP address ranges. Verify there are no overlapping IP addresses and ensure IP ranges use Classless Inter-Domain Routing (CIDR) notation. For example, 172.16.0.0, 192.168.1.0/27. No value is assigned by default'
`--iPAddressWACTokenLifetime [iPAddressWACTokenLifetime]`|Sets IP Address WAC token lifetime'
`--useFindPeopleInPeoplePicker [useFindPeopleInPeoplePicker]`|Sets use find people in PeoplePicker to true or false. Note: When set to true, users aren\'t able to share with security groups or SharePoint groups. Allowed values true,false
`--defaultSharingLinkType [defaultSharingLinkType]`|Lets administrators choose what type of link appears is selected in the “Get a link” sharing dialog box in OneDrive for Business and SharePoint Online. Allowed values None,Direct,Internal,AnonymousAccess
`--oDBMembersCanShare [oDBMembersCanShare]`|Lets administrators set policy on re-sharing behavior in OneDrive for Business. Allowed values Unspecified,On,Off
`--oDBAccessRequests [oDBAccessRequests]`|Lets administrators set policy on access requests and requests to share in OneDrive for Business. Allowed values Unspecified,On,Off
`--preventExternalUsersFromResharing [preventExternalUsersFromResharing]`|Prevents external users from resharing. Allowed values true,false
`--showPeoplePickerSuggestionsForGuestUsers [showPeoplePickerSuggestionsForGuestUsers]`|Shows people picker suggestions for guest users. Allowed values true,false
`--fileAnonymousLinkType [fileAnonymousLinkType]`|Sets the file anonymous link type to None, View or Edit
`--folderAnonymousLinkType [folderAnonymousLinkType]`|Sets the folder anonymous link type to None, View or Edit
`--notifyOwnersWhenItemsReshared [notifyOwnersWhenItemsReshared]`|When this parameter is set to true and another user re-shares a document from a user\'s OneDrive for Business, the OneDrive for Business owner is notified by email. For additional information about how to configure notifications for external sharing, see Configure notifications for external sharing for OneDrive for Business. Allowed values true,false
`--notifyOwnersWhenInvitationsAccepted [notifyOwnersWhenInvitationsAccepted]`|When this parameter is set to true and when an external user accepts an invitation to a resource in a user\'s OneDrive for Business, the OneDrive for Business owner is notified by email. For additional information about how to configure notifications for external sharing, see Configure notifications for external sharing for OneDrive for Business. Allowed values true,false
`--notificationsInOneDriveForBusinessEnabled [notificationsInOneDriveForBusinessEnabled]`|Enables or disables notifications in OneDrive for business. Allowed values true,false
`--notificationsInSharePointEnabled [notificationsInSharePointEnabled]`|Enables or disables notifications in SharePoint. Allowed values true,false
`--ownerAnonymousNotification [ownerAnonymousNotification]`|Enables or disables owner anonymous notification. Allowed values true,false
`--commentsOnSitePagesDisabled [commentsOnSitePagesDisabled]`|Enables or disables comments on site pages. Allowed values true,false
`--socialBarOnSitePagesDisabled [socialBarOnSitePagesDisabled]`|Enables or disables social bar on site pages. Allowed values true,false
`--orphanedPersonalSitesRetentionPeriod [orphanedPersonalSitesRetentionPeriod]`|Specifies the number of days after a user\'s Active Directory account is deleted that their OneDrive for Business content will be deleted. The value range is in days, between 30 and 3650. The default value is 30'
`--disallowInfectedFileDownload [disallowInfectedFileDownload]`|Prevents the Download button from being displayed on the Virus Found warning page. Allowed values true,false
`--defaultLinkPermission [defaultLinkPermission]`|Choose the dafault permission that is selected when users share. This applies to anonymous access, internal and direct links. Allowed values None,View,Edit
`--conditionalAccessPolicy [conditionalAccessPolicy]`|Configures conditional access policy. Allowed values AllowFullAccess,AllowLimitedAccess,BlockAccess
`--allowDownloadingNonWebViewableFiles [allowDownloadingNonWebViewableFiles]`|Allows downloading non web viewable files. The Allowed values true,false
`--allowEditing [allowEditing]`|Allows editing. Allowed values true,false
`--applyAppEnforcedRestrictionsToAdHocRecipients [applyAppEnforcedRestrictionsToAdHocRecipients]`|Applies app enforced restrictions to AdHoc recipients. Allowed values true,false
`--filePickerExternalImageSearchEnabled [filePickerExternalImageSearchEnabled]`|Enables file picker external image search. Allowed values true,false
`--emailAttestationRequired [emailAttestationRequired]`|Sets email attestation to required. Allowed values true,false
`--emailAttestationReAuthDays [emailAttestationReAuthDays]`|Sets email attestation re-auth days'
`--hideDefaultThemes [hideDefaultThemes]`|Defines if the default themes are visible or hidden. Allowed values true,false
`--blockAccessOnUnmanagedDevices [blockAccessOnUnmanagedDevices]`|Blocks access on unmanaged devices. Allowed values true,false
`--allowLimitedAccessOnUnmanagedDevices [allowLimitedAccessOnUnmanagedDevices]`|Allows limited access on unmanaged devices blocks. Allowed values true,false
`--blockDownloadOfAllFilesForGuests [blockDownloadOfAllFilesForGuests]`|Blocks download of all files for guests. Allowed values true,false
`--blockDownloadOfAllFilesOnUnmanagedDevices [blockDownloadOfAllFilesOnUnmanagedDevices]`|Blocks download of all files on unmanaged devices. Allowed values true,false
`--blockDownloadOfViewableFilesForGuests [blockDownloadOfViewableFilesForGuests]`|Blocks download of viewable files for guests. Allowed values true,false
`--blockDownloadOfViewableFilesOnUnmanagedDevices [blockDownloadOfViewableFilesOnUnmanagedDevices]`|Blocks download of viewable files on unmanaged devices. Allowed values true,false
`--blockMacSync [blockMacSync]`|Blocks Mac sync. Allowed values true,false
`--disableReportProblemDialog [disableReportProblemDialog]`|Disables report problem dialog. Allowed values true,false
`--displayNamesOfFileViewers [displayNamesOfFileViewers]`|Displayes names of file viewers. Allowed values true,false
`--enableMinimumVersionRequirement [enableMinimumVersionRequirement]`|Enables minimum version requirement. Allowed values true,false
`--hideSyncButtonOnODB [hideSyncButtonOnODB]`|Hides the sync button on One Drive for Business. Allowed values true,false
`--isUnmanagedSyncClientForTenantRestricted [isUnmanagedSyncClientForTenantRestricted]`|Is unmanaged sync client for tenant restricted. Allowed values true,false
`--limitedAccessFileType [limitedAccessFileType]`|Allows users to preview only Office files in the browser. This option increases security but may be a barrier to user productivity. Allowed values OfficeOnlineFilesOnly,WebPreviewableFiles,OtherFiles
`--optOutOfGrooveBlock [optOutOfGrooveBlock]`|Opts out of the groove block. Allowed values true,false
`--optOutOfGrooveSoftBlock [optOutOfGrooveSoftBlock]`|Opts out of Groove soft block. Allowed values true,false
`--orgNewsSiteUrl [orgNewsSiteUrl]`|Organization news site url'
`--permissiveBrowserFileHandlingOverride [permissiveBrowserFileHandlingOverride]`|Permissive browser fileHandling override. Allowed values true,false
`--showNGSCDialogForSyncOnODB [showNGSCDialogForSyncOnODB]`|Show NGSC dialog for sync on OneDrive for Business. Allowed values true,false
`--specialCharactersStateInFileFolderNames [specialCharactersStateInFileFolderNames]`|Sets the special characters state in file and folder names in SharePoint and OneDrive for Business. Allowed values NoPreference,Allowed,Disallowed
`--syncPrivacyProfileProperties [syncPrivacyProfileProperties]`|Syncs privacy profile properties. Allowed values true,false
`--excludedFileExtensionsForSyncClient [excludedFileExtensionsForSyncClient]`|Excluded file extensions for sync client. Array of strings split by comma (\',\')'
`--allowedDomainListForSyncClient [allowedDomainListForSyncClient]`|Sets allowed domain list for sync client. Array of GUIDs split by comma (\',\'). Example:c9b1909e-901a-0000-2cdb-e91c3f46320a,c9b1909e-901a-0000-2cdb-e91c3f463201'
`--disabledWebPartIds [disabledWebPartIds]`|Sets disabled web part Ids. Array of GUIDs split by comma (\',\'). Example:c9b1909e-901a-0000-2cdb-e91c3f46320a,c9b1909e-901a-0000-2cdb-e91c3f463201'

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Examples

Sets single tenant global setting

```sh
spo tenant settings set --userVoiceForFeedbackEnabled true
```

Sets multiple tenant global settings at once

```sh
spo tenant settings set --userVoiceForFeedbackEnabled true --hideSyncButtonOnODB true --allowedDomainListForSyncClient c9b1909e-901a-0000-2cdb-e91c3f46320a,c9b1909e-901a-0000-2cdb-e91c3f463201
```

## More information

- PnP PowerShell Set-PnPTenant:
[https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnptenant?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnptenant?view=sharepoint-ps)
- SharePoint Online Set-SPOTenant:
[https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenant?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenant?view=sharepoint-ps)
- SharePoint Online Set-SPOTenantCdnEnabled:
[https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenantcdnenabled?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenantcdnenabled?view=sharepoint-ps)
- SharePoint Online Set-SPOTenantSyncClientRestriction: [https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenantsyncclientrestriction?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenantsyncclientrestriction?view=sharepoint-ps)