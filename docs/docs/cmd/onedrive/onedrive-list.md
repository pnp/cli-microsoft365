# onedrive list

Retrieves a list of OneDrive sites

## Usage

```sh
m365 onedrive list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Retrieves a list of OneDrive sites from the tenant.

```sh
m365 onedrive list
```

## Response

=== "JSON"

```json
[
  {
    "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
    "_ObjectIdentity_": "9bc372a0-7035-5000-5787-746234417c24|908bed80-a04a-4433-b4a0-883d9847d110:92ab9c96-469b-4d78-8b8c-961c4df9356b\\\nSiteProperties\\\nhttps%3a%2f%2fcontoso-my.sharepoint.com%2fpersonal%2fuser1_contoso_onmicrosoft_com",
    "AllowDownloadingNonWebViewableFiles": false,
    "AllowEditing": false,
    "AllowSelfServiceUpgrade": true,
    "AnonymousLinkExpirationInDays": 0,
    "AuthContextStrength": null,
    "AuthenticationContextName": null,
    "AverageResourceUsage": 0,
    "BlockDownloadLinksFileType": 0,
    "BlockDownloadMicrosoft365GroupIds": null,
    "BlockDownloadPolicy": false,
    "CommentsOnSitePagesDisabled": false,
    "CompatibilityLevel": 15,
    "ConditionalAccessPolicy": 0,
    "CurrentResourceUsage": 0,
    "DefaultLinkPermission": 0,
    "DefaultLinkToExistingAccess": false,
    "DefaultLinkToExistingAccessReset": false,
    "DefaultShareLinkRole": 0,
    "DefaultShareLinkScope": 0,
    "DefaultSharingLinkType": 0,
    "DenyAddAndCustomizePages": 2,
    "Description": null,
    "DisableAppViews": 0,
    "DisableCompanyWideSharingLinks": 0,
    "DisableFlows": 0,
    "ExcludeBlockDownloadPolicySiteOwners": false,
    "ExcludedBlockDownloadGroupIds": [],
    "ExternalUserExpirationInDays": 0,
    "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "GroupOwnerLoginName": null,
    "HasHolds": false,
    "HubSiteId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "IBMode": null,
    "IBSegments": [],
    "IBSegmentsToAdd": null,
    "IBSegmentsToRemove": null,
    "IsGroupOwnerSiteAdmin": false,
    "IsHubSite": false,
    "IsTeamsChannelConnected": false,
    "IsTeamsConnected": false,
    "LastContentModifiedDate": "/Date(2022,9,27,15,22,10,240)/",
    "Lcid": 1033,
    "LimitedAccessFileType": 0,
    "LockIssue": null,
    "LockState": "Unlock",
    "LoopDefaultSharingLinkRole": 0,
    "LoopDefaultSharingLinkScope": 0,
    "LoopOverrideSharingCapability": false,
    "LoopSharingCapability": 0,
    "MediaTranscription": 0,
    "OverrideBlockUserInfoVisibility": 0,
    "OverrideSharingCapability": false,
    "OverrideTenantAnonymousLinkExpirationPolicy": false,
    "OverrideTenantExternalUserExpirationPolicy": false,
    "Owner": "petkir@contoso.onmicrosoft.com",
    "OwnerEmail": null,
    "OwnerLoginName": null,
    "OwnerName": null,
    "PWAEnabled": 1,
    "ReadOnlyAccessPolicy": false,
    "ReadOnlyForUnmanagedDevices": false,
    "RelatedGroupId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "RequestFilesLinkEnabled": false,
    "RequestFilesLinkExpirationInDays": 0,
    "RestrictedAccessControl": false,
    "RestrictedToRegion": 3,
    "SandboxedCodeActivationCapability": 0,
    "SensitivityLabel": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "SensitivityLabel2": null,
    "SetOwnerWithoutUpdatingSecondaryAdmin": false,
    "SharingAllowedDomainList": null,
    "SharingBlockedDomainList": null,
    "SharingCapability": 2,
    "SharingDomainRestrictionMode": 0,
    "SharingLockDownCanBeCleared": false,
    "SharingLockDownEnabled": false,
    "ShowPeoplePickerSuggestionsForGuestUsers": false,
    "SiteDefinedSharingCapability": 2,
    "SocialBarOnSitePagesDisabled": false,
    "Status": "Active",
    "StorageMaximumLevel": 1048576,
    "StorageQuotaType": null,
    "StorageUsage": 3,
    "StorageWarningLevel": 943718,
    "TeamsChannelType": 0,
    "Template": "SPSPERS#10",
    "TimeZoneId": 13,
    "Title": "Demo User1",
    "TitleTranslations": null,
    "Url": "https://contoso-my.sharepoint.com/personal/user1_contoso_onmicrosoft_com",
    "UserCodeMaximumLevel": 300,
    "UserCodeWarningLevel": 200,
    "WebsCount": 0
  },
  {
    "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties",
    "_ObjectIdentity_": "eac372a0-9026-5000-60c2-f61e1697bf1b|908bed80-a04a-4433-b4a0-883d9847d110:92ab9c96-469b-4d78-8b8c-961c4df9356b\\\nSiteProperties\\\nhttps%3a%2f%2fcontoso-my.sharepoint.com%2fpersonal%2fuser2_contoso_onmicrosoft_com",
    "AllowDownloadingNonWebViewableFiles": false,
    "AllowEditing": false,
    "AllowSelfServiceUpgrade": true,
    "AnonymousLinkExpirationInDays": 0,
    "AuthContextStrength": null,
    "AuthenticationContextName": null,
    "AverageResourceUsage": 0,
    "BlockDownloadLinksFileType": 0,
    "BlockDownloadMicrosoft365GroupIds": null,
    "BlockDownloadPolicy": false,
    "CommentsOnSitePagesDisabled": false,
    "CompatibilityLevel": 15,
    "ConditionalAccessPolicy": 0,
    "CurrentResourceUsage": 0,
    "DefaultLinkPermission": 0,
    "DefaultLinkToExistingAccess": false,
    "DefaultLinkToExistingAccessReset": false,
    "DefaultShareLinkRole": 0,
    "DefaultShareLinkScope": 0,
    "DefaultSharingLinkType": 0,
    "DenyAddAndCustomizePages": 2,
    "Description": null,
    "DisableAppViews": 0,
    "DisableCompanyWideSharingLinks": 0,
    "DisableFlows": 0,
    "ExcludeBlockDownloadPolicySiteOwners": false,
    "ExcludedBlockDownloadGroupIds": [],
    "ExternalUserExpirationInDays": 0,
    "GroupId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "GroupOwnerLoginName": null,
    "HasHolds": false,
    "HubSiteId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "IBMode": null,
    "IBSegments": [],
    "IBSegmentsToAdd": null,
    "IBSegmentsToRemove": null,
    "IsGroupOwnerSiteAdmin": false,
    "IsHubSite": false,
    "IsTeamsChannelConnected": false,
    "IsTeamsConnected": false,
    "LastContentModifiedDate": "/Date(2021,1,23,0,9,1,407)/",
    "Lcid": 1033,
    "LimitedAccessFileType": 0,
    "LockIssue": null,
    "LockState": "Unlock",
    "LoopDefaultSharingLinkRole": 0,
    "LoopDefaultSharingLinkScope": 0,
    "LoopOverrideSharingCapability": false,
    "LoopSharingCapability": 0,
    "MediaTranscription": 0,
    "OverrideBlockUserInfoVisibility": 0,
    "OverrideSharingCapability": false,
    "OverrideTenantAnonymousLinkExpirationPolicy": false,
    "OverrideTenantExternalUserExpirationPolicy": false,
    "Owner": "user2@contoso.onmicrosoft.com",
    "OwnerEmail": null,
    "OwnerLoginName": null,
    "OwnerName": null,
    "PWAEnabled": 1,
    "ReadOnlyAccessPolicy": false,
    "ReadOnlyForUnmanagedDevices": false,
    "RelatedGroupId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "RequestFilesLinkEnabled": false,
    "RequestFilesLinkExpirationInDays": 0,
    "RestrictedAccessControl": false,
    "RestrictedToRegion": 3,
    "SandboxedCodeActivationCapability": 0,
    "SensitivityLabel": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "SensitivityLabel2": null,
    "SetOwnerWithoutUpdatingSecondaryAdmin": false,
    "SharingAllowedDomainList": null,
    "SharingBlockedDomainList": null,
    "SharingCapability": 2,
    "SharingDomainRestrictionMode": 0,
    "SharingLockDownCanBeCleared": false,
    "SharingLockDownEnabled": false,
    "ShowPeoplePickerSuggestionsForGuestUsers": false,
    "SiteDefinedSharingCapability": 2,
    "SocialBarOnSitePagesDisabled": false,
    "Status": "Active",
    "StorageMaximumLevel": 1048576,
    "StorageQuotaType": null,
    "StorageUsage": 1,
    "StorageWarningLevel": 943718,
    "TeamsChannelType": 0,
    "Template": "SPSPERS#10",
    "TimeZoneId": 13,
    "Title": "Demo User2",
    "TitleTranslations": null,
    "Url": "https://contoso-my.sharepoint.com/personal/user2_contoso_onmicrosoft_com",
    "UserCodeMaximumLevel": 300,
    "UserCodeWarningLevel": 200,
    "WebsCount": 0
  }
]
```

=== "Text"

    ``` text
    Title                 Url
    --------------------  ---------------------------------------------------------------------------------
    Demo User1  https://contoso-my.sharepoint.com/personal/user1_contoso_onmicrosoft_com
    Demo User2  https://contoso-my.sharepoint.com/personal/user2_contoso_onmicrosoft_com
    ```

=== "CSV"

    ``` text
    Title,Url
    Demo User2,https://contoso-my.sharepoint.com/personal/user2_contoso_onmicrosoft_com
    ```
