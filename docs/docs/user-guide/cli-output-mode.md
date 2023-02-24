---
title: CLI output mode
sidebar_position: 5
---

# CLI for Microsoft 365 output mode

CLI for Microsoft 365 commands can present their output either as plain-text, JSON, or as CSV. Following is information on these three output modes along with information when to use which.

## Choose the command output mode

All commands in CLI for Microsoft 365 can present their output as plain-text, JSON, CSV or Markdown. By default, all commands use the JSON output mode, but by setting the `--output`, or `-o` for short, option to `text`, you can change the output mode for that command to text. By setting the output option to `csv`, you can change the output mode for that command to CSV. By setting the output option to `md`, you can change the output mode for that command to Markdown.

## JSON output mode

By default, all commands in CLI for Microsoft 365 present their output as JSON strings. This is invaluable when building scripts using the CLI, where the output of one command, has to be processed by another command.

### Simple values

Simple values in JSON output, are returned as-is. For example, if the Microsoft 365 Public CDN is enabled on the currently connected tenant, executing the `spo cdn get` command, will return `true`:

```sh
$ m365 spo cdn get --output json
true
```

### Objects

If the command returns an object, that object will be formatted as a JSON string. For example, getting information about a specific app, will return output similar to:

```sh
$ m365 spo app get -i e6362993-d4fd-4c5a-8254-fd095a7291ad
{
  "AppCatalogVersion": "1.0.0.0",
  "CanUpgrade": false,
  "CurrentVersionDeployed": false,
  "Deployed": false,
  "ID": "e6362993-d4fd-4c5a-8254-fd095a7291ad",
  "InstalledVersion": "",
  "IsClientSideSolution": true,
  "Title": "spfx-140-online-client-side-solution"
}
```

### Arrays

If the command returns information about multiple objects, the command will return a JSON array with each array item representing one object. For example, getting the list of available apps, will return output similar to:

```sh
$ m365 spo app list --output json
[
  {
    "AppCatalogVersion": "1.0.0.0",
    "CanUpgrade": false,
    "CurrentVersionDeployed": false,
    "Deployed": false,
    "ID": "e6362993-d4fd-4c5a-8254-fd095a7291ad",
    "InstalledVersion": "",
    "IsClientSideSolution": true,
    "Title": "spfx-140-online-client-side-solution"
  },
  {
    "AppCatalogVersion": "1.0.0.0",
    "CanUpgrade": false,
    "CurrentVersionDeployed": false,
    "Deployed": false,
    "ID": "5ae74650-b00b-46a9-925f-9c9bd70a0cb6",
    "InstalledVersion": "",
    "IsClientSideSolution": true,
    "Title": "spfx-134-client-side-solution"
  }
]
```

Even if the array contains only one item, for consistency it will be returned as a one-element JSON array:

```sh
$ m365 spo app list --output json
[
  {
    "AppCatalogVersion": "1.0.0.0",
    "CanUpgrade": false,
    "CurrentVersionDeployed": false,
    "Deployed": false,
    "ID": "e6362993-d4fd-4c5a-8254-fd095a7291ad",
    "InstalledVersion": "",
    "IsClientSideSolution": true,
    "Title": "spfx-140-online-client-side-solution"
  }
]
```

!!! tip
    Some `list` commands return different output in text and JSON mode. For readability, in the text mode they only include a few properties, so that the output can be formatted as a table and will fit on the screen. In JSON mode however, they will include all available properties so that it's possible to process the full set of information about the particular object. For more details, refer to the help of the particular command.

### Verbose and debug output in JSON mode

When executing commands in JSON output mode with the `--verbose` or `--debug` flag, the verbose and debug logging statements will be also formatted as JSON and will be added to the output. When processing the command output, you would have to determine yourself which of the returned JSON objects represents the actual command result and which are additional verbose and debug logging statements.

## Text output mode

Optionally, you can have all CLI for Microsoft 365 commands return their output as plain-text. Depending on the command output, the value is presented as-is or formatted for readability.

### Simple values

If the command output is a simple value, such as a number, boolean or a string, the value is returned as-is. For example, if the Microsoft 365 Public CDN is enabled on the currently connected tenant, executing the `spo cdn get` command, will return `true`:

```sh
$ m365 spo cdn get --output text
true
```

### Objects

If the command returns information about an object such as a site, list or an app, that contains a number of properties, the output in text mode is formatted as key-value pairs. For example, getting information about a specific app, will return output similar to:

```sh
$ m365 spo app get -i e6362993-d4fd-4c5a-8254-fd095a7291ad --output text
AppCatalogVersion     : 1.0.0.0
CanUpgrade            : false
CurrentVersionDeployed: false
Deployed              : false
ID                    : e6362993-d4fd-4c5a-8254-fd095a7291ad
InstalledVersion      :
IsClientSideSolution  : true
Title                 : spfx-140-online-client-side-solution
```

### Arrays

If the command returns information about multiple objects, the output is formatted as a table. For example, getting the list of available apps, will return output similar to:

```sh
$ m365 spo app list --output text
Title                                 ID                                    Deployed  AppCatalogVersion
------------------------------------  ------------------------------------  --------  -----------------
spfx-140-online-client-side-solution  e6362993-d4fd-4c5a-8254-fd095a7291ad  false     1.0.0.0
spfx-134-client-side-solution         5ae74650-b00b-46a9-925f-9c9bd70a0cb6  false     1.0.0.0
```

If only one app is returned, it will be displayed as key-value pairs:

```sh
$ m365 spo app list --output text
AppCatalogVersion: 1.0.0.0
Deployed         : false
ID               : e6362993-d4fd-4c5a-8254-fd095a7291ad
Title            : spfx-140-online-client-side-solution
```

## CSV output mode

Optionally, you can have all CLI for Microsoft 365 commands return their output as comma-separated values. Depending on the command output, the value is presented as-is or formatted for readability.

### Simple values

If the command output is a simple value, such as a number, boolean or a string, the value is returned as-is. For example, if the Microsoft 365 Public CDN is enabled on the currently connected tenant, executing the `spo cdn get` command, will return `true`:

```sh
$ m365 spo cdn get --output csv
true
```

### Objects

If the command returns information about an object such as a site, list or an app, that contains a number of properties, the output in CSV mode is formatted as comma-separated values. For example, getting information about a specific app, will return output similar to the following:

```sh
$ m365 spo app get -i e6362993-d4fd-4c5a-8254-fd095a7291ad --output csv
AppCatalogVersion,CanUpgrade,CurrentVersionDeployed,Deployed,ID,InstalledVersion,IsClientSideSolution,Title
1.0.0.0,false,false,false,e6362993-d4fd-4c5a-8254-fd095a7291ad,,true,spfx-140-online-client-side-solution
```

### Arrays

If the command returns information about multiple objects, the output is formatted as multiple lines of comma-separated values. For example, getting the list of available apps will return output similar to the following:

```sh
$ m365 spo app list --output csv
Title,ID,Deployed,AppCatalogVersion
spfx-140-online-client-side-solution,e6362993-d4fd-4c5a-8254-fd095a7291ad,false,1.0.0.0
spfx-134-client-side-solution,5ae74650-b00b-46a9-925f-9c9bd70a0cb6,false,1.0.0.0
```

## Markdown output mode

Using the Markdown output mode is convenient if you need to create documentation for your Microsoft 365 tenant.

!!! tip
    When using the Markdown output, you'll typically want to store the output in a file or in the clipboard. To redirect the output to a file, execute `m365 spo site list --output markdown > sites.md`. To copy the output to the clipboard, on macOS execute `m365 spo site list --output markdown | pbcopy`, and on Windows execute `m365 spo site list --output markdown | clip`.

### Simple values

If the command output is a simple value, such as a number, boolean or a string, the value is returned as-is. For example, if the Microsoft 365 Public CDN is enabled on the currently connected tenant, executing the `spo cdn get` command, will return `true`:

```sh
$ m365 spo cdn get --output md
true
```

### Objects

If the command returns information about an object such as a site, a list or an app, the output in Markdown mode is formatted as a simple report.

```sh
$ m365 spo app get --id c1c89e1f-2332-41f6-aa85-bbb1677262c1 --output md
# spo app get --id "c1c89e1f-2332-41f6-aa85-bbb1677262c1"

Date: 04/12/2022

## spfx-teams-client-side-solution (c1c89e1f-2332-41f6-aa85-bbb1677262c1)

Property | Value
---------|-------
AadAppId | 00000000-0000-0000-0000-000000000000
AadPermissions | null
AppCatalogVersion | 1.0.0.0
CanUpgrade | false
CDNLocation | SharePoint Online
ContainsTenantWideExtension | false
CurrentVersionDeployed | true
Deployed | true
ErrorMessage | No errors.
ID | c1c89e1f-2332-41f6-aa85-bbb1677262c1
InstalledVersion | 
IsClientSideSolution | true
IsEnabled | true
IsPackageDefaultSkipFeatureDeployment | true
IsValidAppPackage | true
ProductId | 3d7d71e9-3bdc-4706-a1d8-59da855f4064
ShortDescription | spfx-teams description
SkipDeploymentFeature | true
StoreAssetId | 
ThumbnailUrl | 
Title | spfx-teams-client-side-solution
```

The report consists of a title section, which shows the information about the executed command and the date when the command was executed.

```markdown
# spo app get --id "c1c89e1f-2332-41f6-aa85-bbb1677262c1"

Date: 04/12/2022

...
```

Then, the report shows information for the retrieved object. The object-specific information starts with a heading, which contains the object's display name and ID.

```markdown
# spo app get --id "c1c89e1f-2332-41f6-aa85-bbb1677262c1"

Date: 04/12/2022

## spfx-teams-client-side-solution (c1c89e1f-2332-41f6-aa85-bbb1677262c1)

...
```

CLI for Microsoft 365 tries to retrieve the object's display name from the following properties in the following order: `title`, `Title`, `displayName`, `DisplayName`, `name`, and `Name`. If the object doesn't have any of these properties, the display name will be  `undefined`.

The display name is followed by the object's ID in parentheses, which CLI tries to resolve from the following properties in the following order: `id`, `Id`, `ID`, `uniqueId`, `UniqueId`, `objectId`, `ObjectId`, `url`, `Url`, `URL`. If the object doesn't have any of these properties, the ID will be displayed as `undefined`.

The heading is followed by a table showing all retrieved object's properties and their values:

```markdown
# spo app get --id "c1c89e1f-2332-41f6-aa85-bbb1677262c1"

Date: 04/12/2022

## spfx-teams-client-side-solution (c1c89e1f-2332-41f6-aa85-bbb1677262c1)

Property | Value
---------|-------
AadAppId | 00000000-0000-0000-0000-000000000000
AadPermissions | null
AppCatalogVersion | 1.0.0.0
CanUpgrade | false
CDNLocation | SharePoint Online
...
```

If the value of a property is an object, it will be JSON-serialized and displayed as a string, for example see the value of the `CurrentChangeToken` property for a site:

```sh
$ m365 spo site get --url /sites/Retail --output md
# spo site get --url "https://contoso.sharepoint.com/sites/Retail"

Date: 04/12/2022

## undefined (4ecc3c2d-4484-4e44-a154-071e6ad711a9)

Property | Value
---------|-------
AllowCreateDeclarativeWorkflow | false
AllowDesigner | true
AllowMasterPageEditing | false
AllowRevertFromTemplate | false
AllowSaveDeclarativeWorkflowAsTemplate | false
AllowSavePublishDeclarativeWorkflow | false
AllowSelfServiceUpgrade | true
AllowSelfServiceUpgradeEvaluation | true
AuditLogTrimmingRetention | 90
ChannelGroupId | 00000000-0000-0000-0000-000000000000
Classification | 
CompatibilityLevel | 15
CurrentChangeToken | {"StringValue":"1;1;4ecc3c2d-4484-4e44-a154-071e6ad711a9;638057455441300000;64815409"}
DisableAppViews | false
DisableCompanyWideSharingLinks | false
DisableFlows | false
ExternalSharingTipsEnabled | false
GeoLocation | EUR
GroupId | 336d4890-fe42-4daa-a53a-8338372a0e59
HubSiteId | 00000000-0000-0000-0000-000000000000
Id | 4ecc3c2d-4484-4e44-a154-071e6ad711a9
SensitivityLabelId | null
SensitivityLabel | 00000000-0000-0000-0000-000000000000
IsHubSite | false
LockIssue | null
MaxItemsPerThrottledOperation | 5000
MediaTranscriptionDisabled | false
NeedsB2BUpgrade | false
ResourcePath | {"DecodedUrl":"https://contoso.sharepoint.com/sites/Retail"}
PrimaryUri | https://contoso.sharepoint.com/sites/Retail
ReadOnly | false
RequiredDesignerVersion | 15.0.0.0
SandboxedCodeActivationCapability | 2
ServerRelativeUrl | /sites/Retail
ShareByEmailEnabled | true
ShareByLinkEnabled | false
ShowUrlStructure | false
TrimAuditLog | true
UIVersionConfigurationEnabled | false
UpgradeReminderDate | 1899-12-30T00:00:00
UpgradeScheduled | false
UpgradeScheduledDate | 1753-01-01T00:00:00
Upgrading | false
Url | https://contoso.sharepoint.com/sites/Retail
WriteLocked | false
```

### Arrays

If the command returns information about multiple objects, the output is formatted as a report, where each object is displayed in a separate section, for example:

```sh
$ m365 spo site list --output md
# spo site list 

Date: 04/12/2022

## Retail (https://contoso.sharepoint.com/sites/Retail)

Property | Value
---------|-------
\_ObjectType\_ | Microsoft.Online.SharePoint.TenantAdministration.SiteProperties
\_ObjectIdentity\_ | 5eed7ea0-c0ab-5000-b47b-32ee0c3bc3d5\|908bed80-a04a-4433-b4a0-883d9847d110:02b03c8c-a55c-4f23-9285-d5bd8f81979a<br>SiteProperties<br>https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fRetail
AllowDownloadingNonWebViewableFiles | false
AllowEditing | false
AllowSelfServiceUpgrade | true
AnonymousLinkExpirationInDays | 0
AuthContextStrength | null
AuthenticationContextName | null
AverageResourceUsage | 0
BlockDownloadLinksFileType | 0
BlockDownloadMicrosoft365GroupIds | null
BlockDownloadPolicy | false
CommentsOnSitePagesDisabled | false
CompatibilityLevel | 15
ConditionalAccessPolicy | 0
CurrentResourceUsage | 0
DefaultLinkPermission | 0
DefaultLinkToExistingAccess | false
DefaultLinkToExistingAccessReset | false
DefaultShareLinkRole | 0
DefaultShareLinkScope | 0
DefaultSharingLinkType | 0
DenyAddAndCustomizePages | 2
Description | null
DisableAppViews | 0
DisableCompanyWideSharingLinks | 0
DisableFlows | 0
ExcludeBlockDownloadPolicySiteOwners | false
ExcludedBlockDownloadGroupIds | []
ExternalUserExpirationInDays | 0
GroupId | /Guid(336d4890-fe42-4daa-a53a-8338372a0e59)/
GroupOwnerLoginName | null
HasHolds | false
HubSiteId | /Guid(00000000-0000-0000-0000-000000000000)/
IBMode | null
IBSegments | []
IBSegmentsToAdd | null
IBSegmentsToRemove | null
IsGroupOwnerSiteAdmin | false
IsHubSite | false
IsTeamsChannelConnected | false
IsTeamsConnected | true
LastContentModifiedDate | /Date(2022,9,18,13,22,1,233)/
Lcid | 1033
LimitedAccessFileType | 0
LockIssue | null
LockState | Unlock
LoopDefaultSharingLinkRole | 0
LoopDefaultSharingLinkScope | 0
LoopOverrideSharingCapability | false
LoopSharingCapability | 0
MediaTranscription | 0
OverrideBlockUserInfoVisibility | 0
OverrideSharingCapability | false
OverrideTenantAnonymousLinkExpirationPolicy | false
OverrideTenantExternalUserExpirationPolicy | false
Owner | 
OwnerEmail | null
OwnerLoginName | null
OwnerName | null
PWAEnabled | 1
ReadOnlyAccessPolicy | false
ReadOnlyForUnmanagedDevices | false
RelatedGroupId | /Guid(336d4890-fe42-4daa-a53a-8338372a0e59)/
RequestFilesLinkEnabled | false
RequestFilesLinkExpirationInDays | 0
RestrictedAccessControl | false
RestrictedToRegion | 3
SandboxedCodeActivationCapability | 0
SensitivityLabel | /Guid(00000000-0000-0000-0000-000000000000)/
SensitivityLabel2 | null
SetOwnerWithoutUpdatingSecondaryAdmin | false
SharingAllowedDomainList | null
SharingBlockedDomainList | null
SharingCapability | 1
SharingDomainRestrictionMode | 0
SharingLockDownCanBeCleared | false
SharingLockDownEnabled | false
ShowPeoplePickerSuggestionsForGuestUsers | false
SiteDefinedSharingCapability | 1
SocialBarOnSitePagesDisabled | false
Status | Active
StorageMaximumLevel | 26214400
StorageQuotaType | null
StorageUsage | 1
StorageWarningLevel | 25574400
TeamsChannelType | 0
Template | GROUP#0
TimeZoneId | 13
Title | Retail
TitleTranslations | null
Url | https://contoso.sharepoint.com/sites/Retail
UserCodeMaximumLevel | 300
UserCodeWarningLevel | 200
WebsCount | 0

## Mark 8 Project Team (https://contoso.sharepoint.com/sites/Mark8ProjectTeam)

Property | Value
---------|-------
\_ObjectType\_ | Microsoft.Online.SharePoint.TenantAdministration.SiteProperties
\_ObjectIdentity\_ | 5eed7ea0-c0ab-5000-b47b-32ee0c3bc3d5\|908bed80-a04a-4433-b4a0-883d9847d110:02b03c8c-a55c-4f23-9285-d5bd8f81979a<br>SiteProperties<br>https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fMark8ProjectTeam
AllowDownloadingNonWebViewableFiles | false
AllowEditing | false
AllowSelfServiceUpgrade | true
AnonymousLinkExpirationInDays | 0
AuthContextStrength | null
AuthenticationContextName | null
AverageResourceUsage | 0
BlockDownloadLinksFileType | 0
BlockDownloadMicrosoft365GroupIds | null
BlockDownloadPolicy | false
CommentsOnSitePagesDisabled | false
CompatibilityLevel | 15
ConditionalAccessPolicy | 0
CurrentResourceUsage | 0
DefaultLinkPermission | 0
DefaultLinkToExistingAccess | false
DefaultLinkToExistingAccessReset | false
DefaultShareLinkRole | 0
DefaultShareLinkScope | 0
DefaultSharingLinkType | 0
DenyAddAndCustomizePages | 2
Description | null
DisableAppViews | 0
DisableCompanyWideSharingLinks | 0
DisableFlows | 0
ExcludeBlockDownloadPolicySiteOwners | false
ExcludedBlockDownloadGroupIds | []
ExternalUserExpirationInDays | 0
GroupId | /Guid(c098b6d1-b1a5-4909-b6ec-ee00aff07b6b)/
GroupOwnerLoginName | null
HasHolds | false
HubSiteId | /Guid(00000000-0000-0000-0000-000000000000)/
IBMode | null
IBSegments | []
IBSegmentsToAdd | null
IBSegmentsToRemove | null
IsGroupOwnerSiteAdmin | false
IsHubSite | false
IsTeamsChannelConnected | false
IsTeamsConnected | true
LastContentModifiedDate | /Date(2022,9,7,12,52,20,347)/
Lcid | 1033
LimitedAccessFileType | 0
LockIssue | null
LockState | Unlock
LoopDefaultSharingLinkRole | 0
LoopDefaultSharingLinkScope | 0
LoopOverrideSharingCapability | false
LoopSharingCapability | 0
MediaTranscription | 0
OverrideBlockUserInfoVisibility | 0
OverrideSharingCapability | false
OverrideTenantAnonymousLinkExpirationPolicy | false
OverrideTenantExternalUserExpirationPolicy | false
Owner | 
OwnerEmail | null
OwnerLoginName | null
OwnerName | null
PWAEnabled | 1
ReadOnlyAccessPolicy | false
ReadOnlyForUnmanagedDevices | false
RelatedGroupId | /Guid(c098b6d1-b1a5-4909-b6ec-ee00aff07b6b)/
RequestFilesLinkEnabled | false
RequestFilesLinkExpirationInDays | 0
RestrictedAccessControl | false
RestrictedToRegion | 3
SandboxedCodeActivationCapability | 0
SensitivityLabel | /Guid(00000000-0000-0000-0000-000000000000)/
SensitivityLabel2 | null
SetOwnerWithoutUpdatingSecondaryAdmin | false
SharingAllowedDomainList | null
SharingBlockedDomainList | null
SharingCapability | 1
SharingDomainRestrictionMode | 0
SharingLockDownCanBeCleared | false
SharingLockDownEnabled | false
ShowPeoplePickerSuggestionsForGuestUsers | false
SiteDefinedSharingCapability | 1
SocialBarOnSitePagesDisabled | false
Status | Active
StorageMaximumLevel | 26214400
StorageQuotaType | null
StorageUsage | 1
StorageWarningLevel | 25574400
TeamsChannelType | 0
Template | GROUP#0
TimeZoneId | 13
Title | Mark 8 Project Team
TitleTranslations | null
Url | https://contoso.sharepoint.com/sites/Mark8ProjectTeam
UserCodeMaximumLevel | 300
UserCodeWarningLevel | 200
WebsCount | 0

...
```

!!! note
    Special Markdown characters, and new line characters in the names and values of properties are escaped so that they're displayed correctly in the table.

## Processing command output with JMESPath

CLI for Microsoft 365 supports filtering, sorting and querying data returned by its commands using [JMESPath](http://jmespath.org/) queries. Queries can be specified using the `--query` option on each command and are applied just before the data retrieved by the command is sent to the console. While you can apply JMESPath queries in all output modes, they are the most powerful in combination with JSON output where the data is unfiltered and the complete data set is sent to output.

For example, you can retrieve the list of all SharePoint site collections in your tenant, by executing:

```sh
$ m365 spo site list --output text
Title                                Url
-----------------------------------  -------------------------------------------------------------------------
Digital Initiative Public Relations  https://contoso.sharepoint.com/sites/DigitalInitiativePublicRelations
Leadership Team                      https://contoso.sharepoint.com/sites/leadership
Mark 8 Project Team                  https://contoso.sharepoint.com/sites/Mark8ProjectTeam
Operations                           https://contoso.sharepoint.com/sites/operations
Retail                               https://contoso.sharepoint.com/sites/Retail
Sales and Marketing                  https://contoso.sharepoint.com/sites/SalesAndMarketing
```

To retrieve information only about sites matching a specific title or URL, you could execute:

```sh
$ m365 spo site list --query "[?Title == 'Retail']" --output text
Title: Retail
Url  : https://contoso.sharepoint.com/sites/Retail
```

To make the output more readable, you could pass it to a JSON processor such as [jq](https://stedolan.github.io/jq/):

```sh
$ m365 spo site list --output json --query "[?Template == 'GROUP#0'].{Title: Title, Url: Url}" | jq
[
  {
    "Title": "Mark 8 Project Team",
    "Url": "https://contoso.sharepoint.com/sites/Mark8ProjectTeam"
  },
  {
    "Title": "Operations",
    "Url": "https://contoso.sharepoint.com/sites/operations"
  },
  {
    "Title": "Digital Initiative Public Relations",
    "Url": "https://contoso.sharepoint.com/sites/DigitalInitiativePublicRelations"
  },
  {
    "Title": "Retail",
    "Url": "https://contoso.sharepoint.com/sites/Retail"
  },
  {
    "Title": "Leadership Team",
    "Url": "https://contoso.sharepoint.com/sites/leadership"
  },
  {
    "Title": "Sales and Marketing",
    "Url": "https://contoso.sharepoint.com/sites/SalesAndMarketing"
  }
]
```

## When to use which output mode

Generally, you will use the text output when interacting with the CLI yourself. When building scripts using the CLI for Microsoft 365, you will use the default JSON output mode, because processing JSON strings is much more convenient and reliable than processing plain-text output.

## Sample script

Using the JSON output mode allows you to build scripts using the CLI for Microsoft 365. The CLI works on any platform, but as there is no common way to work with objects and command output on all platforms and shells, we chose JSON as the format to serialize objects in the CLI for Microsoft 365.

Following, is a sample script, that you could build using the CLI for Microsoft 365 in Bash:

```sh
m365 # get all apps available in the tenant app catalog
apps=$(m365 spo app list --output json)

# get IDs of all apps that are not deployed
notDeployedAppsIds=($(echo $apps | jq -r '.[] | select(.Deployed == false) | {ID} | .[]'))

# deploy all not deployed apps
for appId in $notDeployedAppsIds; do
  m365 spo app deploy -i $appId
done
```

_First, you use the CLI for Microsoft 365 to get the list of all apps from the tenant app catalog using the [spo app list](../cmd/spo/app/app-list.md) command. You set the output type to JSON and store it in a shell variable `apps`. Next, you parse the JSON string using [jq](https://stedolan.github.io/jq/) and get IDs of apps that are not deployed. Finally, for each ID you run the [spo app deploy](../cmd/spo/app/app-deploy.md) CLI for Microsoft 365 command passing the ID as a command argument. Notice, that in the script, both `spo` commands are prepended with `m365` and executed as separate commands directly in the shell._

The same could be accomplished in PowerShell as well:

```powershell
# get all apps available in the tenant app catalog
$apps = m365 spo app list --output json | ConvertFrom-Json

# get all apps that are not yet deployed and deploy them
$apps | ? Deployed -eq $false | % { m365 spo app deploy -i $_.ID }
```

Because PowerShell offers native support for working with JSON strings and objects, the same script written in PowerShell is simpler than the one in Bash. At the end of the day it's up to you to choose if you want to use the CLI for Microsoft 365 in Bash, PowerShell or some other shell. Both Bash and PowerShell are available on multiple platforms, and if you have a team using different platforms, writing scripts using CLI for Microsoft 365 in Bash or PowerShell will let everyone in your team use them.
