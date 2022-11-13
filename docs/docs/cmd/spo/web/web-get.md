# spo web get

Retrieve information about the specified site

## Usage

```sh
m365 spo web get [options]
```

## Options

`-u, --url <url>`
: URL of the site for which to retrieve the information

`--withGroups`
: Set if you want to return associated groups (associatedOwnerGroup, associatedMemberGroup and associatedVisitorGroup) along with other properties

`--withPermissions`
: Set if you want to return associated roles and permissions

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve information about a site

```sh
m365 spo web get --url https://contoso.sharepoint.com/subsite
```

Retrieve information about a site along with associated groups for the web

```sh
m365 spo web get --url https://contoso.sharepoint.com/subsite --withGroups
```

Retrieve information about a site along with the RoleAssignments for the web

```sh
m365 spo web get --url https://contoso.sharepoint.com/subsite --withPermissions
```

## Response

=== "JSON"

    ```json
    {
      "AllowRssFeeds": true,
      "AlternateCssUrl": "",
      "AppInstanceId": "00000000-0000-0000-0000-000000000000",
      "ClassicWelcomePage": null,
      "Configuration": 0,
      "Created": "2022-09-12T18:18:07.253",
      "CurrentChangeToken": {
        "StringValue": "1;2;d8b65bb3-6ca1-4df2-a4be-0efe08af2580;638032553974830000;715586578"
      },
      "CustomMasterUrl": "/_catalogs/masterpage/seattle.master",
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
      "Id": "d8b65bb3-6ca1-4df2-a4be-0efe08af2580",
      "IsEduClass": false,
      "IsEduClassProvisionChecked": false,
      "IsEduClassProvisionPending": false,
      "IsHomepageModernized": false,
      "IsMultilingual": true,
      "IsRevertHomepageLinkHidden": false,
      "Language": 1033,
      "LastItemModifiedDate": "2022-11-05T14:06:21Z",
      "LastItemUserModifiedDate": "2022-10-31T07:29:33Z",
      "LogoAlignment": 0,
      "MasterUrl": "/_catalogs/masterpage/seattle.master",
      "MegaMenuEnabled": true,
      "NavAudienceTargetingEnabled": false,
      "NoCrawl": false,
      "ObjectCacheEnabled": false,
      "OverwriteTranslationsOnChange": false,
      "ResourcePath": {
        "DecodedUrl": "https://contoso.sharepoint.com"
      },
      "QuickLaunchEnabled": true,
      "RecycleBinEnabled": true,
      "SearchScope": 0,
      "ServerRelativeUrl": "/",
      "SiteLogoUrl": null,
      "SyndicationEnabled": true,
      "TenantAdminMembersCanShare": 0,
      "Title": "Communication site",
      "TreeViewEnabled": false,
      "UIVersion": 15,
      "UIVersionConfigurationEnabled": false,
      "Url": "https://contoso.sharepoint.com",
      "WebTemplate": "SITEPAGEPUBLISHING",
      "WelcomePage": "SitePages/Home.aspx"
    }
    ```

=== "Text"

    ```text
    AllowRssFeeds                                       : true
    AlternateCssUrl                                     :
    AppInstanceId                                       : 00000000-0000-0000-0000-000000000000
    ClassicWelcomePage                                  : null
    Configuration                                       : 0
    Created                                             : 2022-09-12T18:18:07.253
    CurrentChangeToken                                  : {"StringValue":"1;2;d8b65bb3-6ca1-4df2-a4be-0efe08af2580;638032554376830000;715586608"}
    CustomMasterUrl                                     : /_catalogs/masterpage/seattle.master
    Description                                         :
    DesignPackageId                                     : 00000000-0000-0000-0000-000000000000
    DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false
    EnableMinimalDownload                               : false
    FooterEmphasis                                      : 0
    FooterEnabled                                       : true
    FooterLayout                                        : 0
    HeaderEmphasis                                      : 0
    HeaderLayout                                        : 0
    HideTitleInHeader                                   : false
    HorizontalQuickLaunch                               : false
    Id                                                  : d8b65bb3-6ca1-4df2-a4be-0efe08af2580
    IsEduClass                                          : false
    IsEduClassProvisionChecked                          : false
    IsEduClassProvisionPending                          : false
    IsHomepageModernized                                : false
    IsMultilingual                                      : true
    IsRevertHomepageLinkHidden                          : false
    Language                                            : 1033
    LastItemModifiedDate                                : 2022-11-05T14:06:21Z
    LastItemUserModifiedDate                            : 2022-10-31T07:29:33Z
    LogoAlignment                                       : 0
    MasterUrl                                           : /_catalogs/masterpage/seattle.master
    MegaMenuEnabled                                     : true
    NavAudienceTargetingEnabled                         : false
    NoCrawl                                             : false
    ObjectCacheEnabled                                  : false
    OverwriteTranslationsOnChange                       : false
    QuickLaunchEnabled                                  : true
    RecycleBinEnabled                                   : true
    ResourcePath                                        : {"DecodedUrl":"https://contoso.sharepoint.com"}
    SearchScope                                         : 0
    ServerRelativeUrl                                   : /
    SiteLogoUrl                                         : null
    SyndicationEnabled                                  : true
    TenantAdminMembersCanShare                          : 0
    Title                                               : Communication site
    TreeViewEnabled                                     : false
    UIVersion                                           : 15
    UIVersionConfigurationEnabled                       : false
    Url                                                 : https://contoso.sharepoint.com
    WebTemplate                                         : SITEPAGEPUBLISHING
    WelcomePage                                         : SitePages/Home.aspx
    ```

=== "CSV"

    ```csv
    AllowRssFeeds,AlternateCssUrl,AppInstanceId,ClassicWelcomePage,Configuration,Created,CurrentChangeToken,CustomMasterUrl,Description,DesignPackageId,DocumentLibraryCalloutOfficeWebAppPreviewersDisabled,EnableMinimalDownload,FooterEmphasis,FooterEnabled,FooterLayout,HeaderEmphasis,HeaderLayout,HideTitleInHeader,HorizontalQuickLaunch,Id,IsEduClass,IsEduClassProvisionChecked,IsEduClassProvisionPending,IsHomepageModernized,IsMultilingual,IsRevertHomepageLinkHidden,Language,LastItemModifiedDate,LastItemUserModifiedDate,LogoAlignment,MasterUrl,MegaMenuEnabled,NavAudienceTargetingEnabled,NoCrawl,ObjectCacheEnabled,OverwriteTranslationsOnChange,ResourcePath,QuickLaunchEnabled,RecycleBinEnabled,SearchScope,ServerRelativeUrl,SiteLogoUrl,SyndicationEnabled,TenantAdminMembersCanShare,Title,TreeViewEnabled,UIVersion,UIVersionConfigurationEnabled,Url,WebTemplate,WelcomePage
    1,,00000000-0000-0000-0000-000000000000,,0,2022-09-12T18:18:07.253,"{""StringValue"":""1;2;d8b65bb3-6ca1-4df2-a4be-0efe08af2580;638032554734300000;715586625""}",/_catalogs/masterpage/seattle.master,,00000000-0000-0000-0000-000000000000,,,0,1,0,0,0,,,d8b65bb3-6ca1-4df2-a4be-0efe08af2580,,,,,1,,1033,2022-11-05T14:06:21Z,2022-10-31T07:29:33Z,0,/_catalogs/masterpage/seattle.master,1,,,,,"{""DecodedUrl"":""https://contoso.sharepoint.com""}",1,1,0,/,,1,0,Communication site,,15,,https://contoso.sharepoint.com,SITEPAGEPUBLISHING,SitePages/Home.aspx
    ```

### `withGroups`, `withPermissions` response

When we make use of the option `withGroups` or `withPermissions` the response will differ. 

## Response

=== "JSON"

    ```json
    {
      "AssociatedMemberGroup": {
        "Id": 5,
        "IsHiddenInUI": false,
        "LoginName": "Communication site Members",
        "Title": "Communication site Members",
        "PrincipalType": 8,
        "AllowMembersEditMembership": true,
        "AllowRequestToJoinLeave": false,
        "AutoAcceptRequestToJoinLeave": false,
        "Description": null,
        "OnlyAllowMembersViewMembership": false,
        "OwnerTitle": "Communication site Owners",
        "RequestToJoinLeaveEmailSetting": ""
      },
      "AssociatedOwnerGroup": {
        "Id": 3,
        "IsHiddenInUI": false,
        "LoginName": "Communication site Owners",
        "Title": "Communication site Owners",
        "PrincipalType": 8,
        "AllowMembersEditMembership": false,
        "AllowRequestToJoinLeave": false,
        "AutoAcceptRequestToJoinLeave": false,
        "Description": null,
        "OnlyAllowMembersViewMembership": false,
        "OwnerTitle": "Communication site Owners",
        "RequestToJoinLeaveEmailSetting": ""
      },
      "AssociatedVisitorGroup": {
        "Id": 4,
        "IsHiddenInUI": false,
        "LoginName": "Communication site Visitors",
        "Title": "Communication site Visitors",
        "PrincipalType": 8,
        "AllowMembersEditMembership": false,
        "AllowRequestToJoinLeave": false,
        "AutoAcceptRequestToJoinLeave": false,
        "Description": null,
        "OnlyAllowMembersViewMembership": false,
        "OwnerTitle": "Communication site Owners",
        "RequestToJoinLeaveEmailSetting": ""
      },
      "AllowRssFeeds": true,
      "AlternateCssUrl": "",
      "AppInstanceId": "00000000-0000-0000-0000-000000000000",
      "ClassicWelcomePage": null,
      "Configuration": 0,
      "Created": "2022-09-12T18:18:07.253",
      "CurrentChangeToken": {
        "StringValue": "1;2;d8b65bb3-6ca1-4df2-a4be-0efe08af2580;638032553974830000;715586578"
      },
      "CustomMasterUrl": "/_catalogs/masterpage/seattle.master",
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
      "Id": "d8b65bb3-6ca1-4df2-a4be-0efe08af2580",
      "IsEduClass": false,
      "IsEduClassProvisionChecked": false,
      "IsEduClassProvisionPending": false,
      "IsHomepageModernized": false,
      "IsMultilingual": true,
      "IsRevertHomepageLinkHidden": false,
      "Language": 1033,
      "LastItemModifiedDate": "2022-11-05T14:06:21Z",
      "LastItemUserModifiedDate": "2022-10-31T07:29:33Z",
      "LogoAlignment": 0,
      "MasterUrl": "/_catalogs/masterpage/seattle.master",
      "MegaMenuEnabled": true,
      "NavAudienceTargetingEnabled": false,
      "NoCrawl": false,
      "ObjectCacheEnabled": false,
      "OverwriteTranslationsOnChange": false,
      "ResourcePath": {
        "DecodedUrl": "https://contoso.sharepoint.com"
      },
      "QuickLaunchEnabled": true,
      "RecycleBinEnabled": true,
      "SearchScope": 0,
      "ServerRelativeUrl": "/",
      "SiteLogoUrl": null,
      "SyndicationEnabled": true,
      "TenantAdminMembersCanShare": 0,
      "Title": "Communication site",
      "TreeViewEnabled": false,
      "UIVersion": 15,
      "UIVersionConfigurationEnabled": false,
      "Url": "https://contoso.sharepoint.com",
      "WebTemplate": "SITEPAGEPUBLISHING",
      "WelcomePage": "SitePages/Home.aspx",
      "RoleAssignments": [
      {
        "Member": {
          "Id": 3,
          "IsHiddenInUI": false,
          "LoginName": "Communication site Owners",
          "Title": "Communication site Owners",
          "PrincipalType": 8,
          "AllowMembersEditMembership": false,
          "AllowRequestToJoinLeave": false,
          "AutoAcceptRequestToJoinLeave": false,
          "Description": null,
          "OnlyAllowMembersViewMembership": false,
          "OwnerTitle": "Communication site Owners",
          "RequestToJoinLeaveEmailSetting": ""
        },
        "RoleDefinitionBindings": [
          {
            "BasePermissions": {
              "High": "2147483647",
              "Low": "4294967295"
            },
            "Description": "Has full control.",
            "Hidden": false,
            "Id": 1073741829,
            "Name": "Full Control",
            "Order": 1,
            "RoleTypeKind": 5,
            "BasePermissionsValue": [
              "ViewListItems",
              "AddListItems",
              "EditListItems",
              "DeleteListItems",
              "ApproveItems",
              "OpenItems",
              "ViewVersions",
              "DeleteVersions",
              "CancelCheckout",
              "ManagePersonalViews",
              "ManageLists",
              "ViewFormPages",
              "AnonymousSearchAccessList",
              "Open",
              "ViewPages",
              "AddAndCustomizePages",
              "ApplyThemeAndBorder",
              "ApplyStyleSheets",
              "ViewUsageData",
              "CreateSSCSite",
              "ManageSubwebs",
              "CreateGroups",
              "ManagePermissions",
              "BrowseDirectories",
              "BrowseUserInfo",
              "AddDelPrivateWebParts",
              "UpdatePersonalWebParts",
              "ManageWeb",
              "AnonymousSearchAccessWebLists",
              "UseClientIntegration",
              "UseRemoteAPIs",
              "ManageAlerts",
              "CreateAlerts",
              "EditMyUserInfo",
              "EnumeratePermissions"
            ],
            "RoleTypeKindValue": "Administrator"
          }
        ],
        "PrincipalId": 3
      }]
    }
    ```

=== "Text"

    ```text
    AllowRssFeeds                                       : true
    AlternateCssUrl                                     :
    AppInstanceId                                       : 00000000-0000-0000-0000-000000000000
    AssociatedMemberGroup                               : {"Id":5,"IsHiddenInUI":false,"LoginName":"Communication site Members","Title":"Communication site Members","PrincipalType":8,"AllowMembersEditMembership":true,"AllowRequestToJoinLeave":false,"AutoAcceptRequestToJoinLeave":false,"Description":null,"OnlyAllowMembersViewMembership":false,"OwnerTitle":"Communication site Owners","RequestToJoinLeaveEmailSetting":""}
    AssociatedOwnerGroup                                : {"Id":3,"IsHiddenInUI":false,"LoginName":"Communication site Owners","Title":"Communication site Owners","PrincipalType":8,"AllowMembersEditMembership":false,"AllowRequestToJoinLeave":false,"AutoAcceptRequestToJoinLeave":false,"Description":null,"OnlyAllowMembersViewMembership":false,"OwnerTitle":"Communication site Owners","RequestToJoinLeaveEmailSetting":""}
    AssociatedVisitorGroup                              : {"Id":4,"IsHiddenInUI":false,"LoginName":"Communication site Visitors","Title":"Communication site Visitors","PrincipalType":8,"AllowMembersEditMembership":false,"AllowRequestToJoinLeave":false,"AutoAcceptRequestToJoinLeave":false,"Description":null,"OnlyAllowMembersViewMembership":false,"OwnerTitle":"Communication site Owners","RequestToJoinLeaveEmailSetting":""}
    ClassicWelcomePage                                  : null
    Configuration                                       : 0
    Created                                             : 2022-09-12T18:18:07.253
    CurrentChangeToken                                  : {"StringValue":"1;2;d8b65bb3-6ca1-4df2-a4be-0efe08af2580;638032554376830000;715586608"}
    CustomMasterUrl                                     : /_catalogs/masterpage/seattle.master
    Description                                         :
    DesignPackageId                                     : 00000000-0000-0000-0000-000000000000
    DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false
    EnableMinimalDownload                               : false
    FooterEmphasis                                      : 0
    FooterEnabled                                       : true
    FooterLayout                                        : 0
    HeaderEmphasis                                      : 0
    HeaderLayout                                        : 0
    HideTitleInHeader                                   : false
    HorizontalQuickLaunch                               : false
    Id                                                  : d8b65bb3-6ca1-4df2-a4be-0efe08af2580
    IsEduClass                                          : false
    IsEduClassProvisionChecked                          : false
    IsEduClassProvisionPending                          : false
    IsHomepageModernized                                : false
    IsMultilingual                                      : true
    IsRevertHomepageLinkHidden                          : false
    Language                                            : 1033
    LastItemModifiedDate                                : 2022-11-05T14:06:21Z
    LastItemUserModifiedDate                            : 2022-10-31T07:29:33Z
    LogoAlignment                                       : 0
    MasterUrl                                           : /_catalogs/masterpage/seattle.master
    MegaMenuEnabled                                     : true
    NavAudienceTargetingEnabled                         : false
    NoCrawl                                             : false
    ObjectCacheEnabled                                  : false
    OverwriteTranslationsOnChange                       : false
    QuickLaunchEnabled                                  : true
    RecycleBinEnabled                                   : true
    ResourcePath                                        : {"DecodedUrl":"https://contoso.sharepoint.com"}
    RoleAssignments                                     : [{"Member":{"Id":3,"IsHiddenInUI":false,"LoginName":"Communication site Owners","Title":"Communication site Owners","PrincipalType":8,"AllowMembersEditMembership":false,"AllowRequestToJoinLeave":false,"AutoAcceptRequestToJoinLeave":false,"Description":null,"OnlyAllowMembersViewMembership":false,"OwnerTitle":"Communication site Owners","RequestToJoinLeaveEmailSetting":""},"RoleDefinitionBindings":[{"BasePermissions":{"High":"2147483647","Low":"4294967295"},"Description":"Has full control.","Hidden":false,"Id":1073741829,"Name":"Full Control","Order":1,"RoleTypeKind":5,"BasePermissionsValue":["ViewListItems","AddListItems","EditListItems","DeleteListItems","ApproveItems","OpenItems","ViewVersions","DeleteVersions","CancelCheckout","ManagePersonalViews","ManageLists","ViewFormPages","AnonymousSearchAccessList","Open","ViewPages","AddAndCustomizePages","ApplyThemeAndBorder","ApplyStyleSheets","ViewUsageData","CreateSSCSite","ManageSubwebs","CreateGroups","ManagePermissions","BrowseDirectories","BrowseUserInfo","AddDelPrivateWebParts","UpdatePersonalWebParts","ManageWeb","AnonymousSearchAccessWebLists","UseClientIntegration","UseRemoteAPIs","ManageAlerts","CreateAlerts","EditMyUserInfo","EnumeratePermissions"],"RoleTypeKindValue":"Administrator"}],"PrincipalId":3}]
    SearchScope                                         : 0
    ServerRelativeUrl                                   : /
    SiteLogoUrl                                         : null
    SyndicationEnabled                                  : true
    TenantAdminMembersCanShare                          : 0
    Title                                               : Communication site
    TreeViewEnabled                                     : false
    UIVersion                                           : 15
    UIVersionConfigurationEnabled                       : false
    Url                                                 : https://contoso.sharepoint.com
    WebTemplate                                         : SITEPAGEPUBLISHING
    WelcomePage                                         : SitePages/Home.aspx
    ```

=== "CSV"

    ```csv
    AssociatedMemberGroup,AssociatedOwnerGroup,AssociatedVisitorGroup,AllowRssFeeds,AlternateCssUrl,AppInstanceId,ClassicWelcomePage,Configuration,Created,CurrentChangeToken,CustomMasterUrl,Description,DesignPackageId,DocumentLibraryCalloutOfficeWebAppPreviewersDisabled,EnableMinimalDownload,FooterEmphasis,FooterEnabled,FooterLayout,HeaderEmphasis,HeaderLayout,HideTitleInHeader,HorizontalQuickLaunch,Id,IsEduClass,IsEduClassProvisionChecked,IsEduClassProvisionPending,IsHomepageModernized,IsMultilingual,IsRevertHomepageLinkHidden,Language,LastItemModifiedDate,LastItemUserModifiedDate,LogoAlignment,MasterUrl,MegaMenuEnabled,NavAudienceTargetingEnabled,NoCrawl,ObjectCacheEnabled,OverwriteTranslationsOnChange,ResourcePath,QuickLaunchEnabled,RecycleBinEnabled,SearchScope,ServerRelativeUrl,SiteLogoUrl,SyndicationEnabled,TenantAdminMembersCanShare,Title,TreeViewEnabled,UIVersion,UIVersionConfigurationEnabled,Url,WebTemplate,WelcomePage,RoleAssignments
    "{""Id"":5,""IsHiddenInUI"":false,""LoginName"":""Communication site Members"",""Title"":""Communication site Members"",""PrincipalType"":8,""AllowMembersEditMembership"":true,""AllowRequestToJoinLeave"":false,""AutoAcceptRequestToJoinLeave"":false,""Description"":null,""OnlyAllowMembersViewMembership"":false,""OwnerTitle"":""Communication site Owners"",""RequestToJoinLeaveEmailSetting"":""""}","{""Id"":3,""IsHiddenInUI"":false,""LoginName"":""Communication site Owners"",""Title"":""Communication site Owners"",""PrincipalType"":8,""AllowMembersEditMembership"":false,""AllowRequestToJoinLeave"":false,""AutoAcceptRequestToJoinLeave"":false,""Description"":null,""OnlyAllowMembersViewMembership"":false,""OwnerTitle"":""Communication site Owners"",""RequestToJoinLeaveEmailSetting"":""""}","{""Id"":4,""IsHiddenInUI"":false,""LoginName"":""Communication site Visitors"",""Title"":""Communication site Visitors"",""PrincipalType"":8,""AllowMembersEditMembership"":false,""AllowRequestToJoinLeave"":false,""AutoAcceptRequestToJoinLeave"":false,""Description"":null,""OnlyAllowMembersViewMembership"":false,""OwnerTitle"":""Communication site Owners"",""RequestToJoinLeaveEmailSetting"":""""}",1,,00000000-0000-0000-0000-000000000000,,0,2022-09-12T18:18:07.253,"{""StringValue"":""1;2;d8b65bb3-6ca1-4df2-a4be-0efe08af2580;638032554734300000;715586625""}",/_catalogs/masterpage/seattle.master,,00000000-0000-0000-0000-000000000000,,,0,1,0,0,0,,,d8b65bb3-6ca1-4df2-a4be-0efe08af2580,,,,,1,,1033,2022-11-05T14:06:21Z,2022-10-31T07:29:33Z,0,/_catalogs/masterpage/seattle.master,1,,,,,"{""DecodedUrl"":""https://contoso.sharepoint.com""}",1,1,0,/,,1,0,Communication site,,15,,https://contoso.sharepoint.com,SITEPAGEPUBLISHING,SitePages/Home.aspx,"[{""Member"":{""Id"":3,""IsHiddenInUI"":false,""LoginName"":""Communication site Owners"",""Title"":""Communication site Owners"",""PrincipalType"":8,""AllowMembersEditMembership"":false,""AllowRequestToJoinLeave"":false,""AutoAcceptRequestToJoinLeave"":false,""Description"":null,""OnlyAllowMembersViewMembership"":false,""OwnerTitle"":""Communication site Owners"",""RequestToJoinLeaveEmailSetting"":""""},""RoleDefinitionBindings"":[{""BasePermissions"":{""High"":""2147483647"",""Low"":""4294967295""},""Description"":""Has full control."",""Hidden"":false,""Id"":1073741829,""Name"":""Full Control"",""Order"":1,""RoleTypeKind"":5,""BasePermissionsValue"":[""ViewListItems"",""AddListItems"",""EditListItems"",""DeleteListItems"",""ApproveItems"",""OpenItems"",""ViewVersions"",""DeleteVersions"",""CancelCheckout"",""ManagePersonalViews"",""ManageLists"",""ViewFormPages"",""AnonymousSearchAccessList"",""Open"",""ViewPages"",""AddAndCustomizePages"",""ApplyThemeAndBorder"",""ApplyStyleSheets"",""ViewUsageData"",""CreateSSCSite"",""ManageSubwebs"",""CreateGroups"",""ManagePermissions"",""BrowseDirectories"",""BrowseUserInfo"",""AddDelPrivateWebParts"",""UpdatePersonalWebParts"",""ManageWeb"",""AnonymousSearchAccessWebLists"",""UseClientIntegration"",""UseRemoteAPIs"",""ManageAlerts"",""CreateAlerts"",""EditMyUserInfo"",""EnumeratePermissions""],""RoleTypeKindValue"":""Administrator""}],""PrincipalId"":3}]"
    ```
