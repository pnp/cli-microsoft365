# spo file get

Gets information about the specified file

## Usage

```sh
m365 spo file get [options]
```

## Options

`-w, --webUrl <webUrl>`
: The URL of the site where the file is located

`-u, --url [url]`
: The server-relative URL of the file to retrieve. Specify either `url` or `id` but not both

`-i, --id [id]`
: The UniqueId (GUID) of the file to retrieve. Specify either `url` or `id` but not both

`--asString`
: Set to retrieve the contents of the specified file as string

`--asListItem`
: Set to retrieve the underlying list item

`--asFile`
: Set to save the file to the path specified in the path option

`-p, --path [path]`
: The local path where to save the retrieved file. Must be specified when the `--asFile` option is used

`--withPermissions`
: Set if you want to return associated roles and permissions

--8<-- "docs/cmd/_global.md"

## Examples

Get file properties for a file with id (UniqueId) parameter located in a site

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Get contents of the file with id (UniqueId) parameter located in a site

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asString
```

Get list item properties for a file with id (UniqueId) parameter located in a site

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asListItem
```

Saves the file with id (UniqueId) parameter located in a site to a local file

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asFile --path /Users/user/documents/SavedAsTest1.docx
```

Return file properties for a file with server-relative url located in a site

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx'
```

Returns a file as string for a file with server-relative url located in a site

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asString
```

Returna the list item properties for a file with the server-relative url located in a site

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asListItem
```

Saves a file with the server-relative url located in a site to a local file

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asFile --path /Users/user/documents/SavedAsTest1.docx
```

Gets the file properties for a file with id (UniqueId) parameter located in a site with permissions

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --withPermissions
```

## Response

=== "JSON"

    ```json
    {
      "CheckInComment": "",
      "CheckOutType": 2,
      "ContentTag": "{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1,1",
      "CustomizedPageStatus": 0,
      "ETag": "\"{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1\"",
      "Exists": true,
      "IrmEnabled": false,
      "Length": "5987",
      "Level": 1,
      "LinkingUri": null,
      "LinkingUrl": "",
      "MajorVersion": 1,
      "MinorVersion": 0,
      "Name": "Test1.docx",
      "ServerRelativeUrl": "/sites/project-x/documents/Test1.docx",
      "TimeCreated": "2022-10-30T10:16:18Z",
      "TimeLastModified": "2022-10-30T10:16:18Z",
      "Title": null,
      "UIVersion": 512,
      "UIVersionLabel": "1.0",
      "UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"
    }
    ```

=== "Text"

    ```text
    CheckInComment      :
    CheckOutType        : 2
    ContentTag          : {03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1,1
    CustomizedPageStatus: 0
    ETag                : "{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1"
    Exists              : true
    IrmEnabled          : false
    Length              : 5987
    Level               : 1
    LinkingUri          : null
    LinkingUrl          :
    MajorVersion        : 1
    MinorVersion        : 0
    Name                : Test1.docx
    ServerRelativeUrl   : /sites/project-x/documents/Test1.docx
    TimeCreated         : 2022-10-30T10:16:18Z
    TimeLastModified    : 2022-10-30T10:16:18Z
    Title               : null
    UIVersion           : 512
    UIVersionLabel      : 1.0
    UniqueId            : b2307a39-e878-458b-bc90-03bc578531d6
    ```

=== "CSV"

    ```csv
    CheckInComment,CheckOutType,ContentTag,CustomizedPageStatus,ETag,Exists,IrmEnabled,Length,Level,LinkingUri,LinkingUrl,MajorVersion,MinorVersion,Name,ServerRelativeUrl,TimeCreated,TimeLastModified,Title,UIVersion,UIVersionLabel,UniqueId
    ,2,"{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1,1",0,"""{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1""",1,,5987,1,,,1,0,Test1.docx,/sites/project-x/documents/Test1.docx,2022-10-30T10:16:18Z,2022-10-30T10:16:18Z,,512,1.0,b2307a39-e878-458b-bc90-03bc578531d6
    ```

### `withPermissions` response

When we make use of the option `withPermissions` the response will differ.

## Response

=== "JSON"

    ```json
    {
      "CheckInComment": "",
      "CheckOutType": 2,
      "ContentTag": "{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1,1",
      "CustomizedPageStatus": 0,
      "ETag": "\"{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1\"",
      "Exists": true,
      "IrmEnabled": false,
      "Length": "5987",
      "Level": 1,
      "LinkingUri": null,
      "LinkingUrl": "",
      "MajorVersion": 1,
      "MinorVersion": 0,
      "Name": "Test1.docx",
      "ServerRelativeUrl": "/sites/project-x/documents/Test1.docx",
      "TimeCreated": "2022-10-30T10:16:18Z",
      "TimeLastModified": "2022-10-30T10:16:18Z",
      "Title": null,
      "UIVersion": 512,
      "UIVersionLabel": "1.0",
      "UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6",
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
        }
      ]
    }
    ```

=== "Text"

    ```text
    CheckInComment      :
    CheckOutType        : 2
    ContentTag          : {03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1,1
    CustomizedPageStatus: 0
    ETag                : "{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1"
    Exists              : true
    IrmEnabled          : false
    Length              : 5987
    Level               : 1
    LinkingUri          : null
    LinkingUrl          :
    MajorVersion        : 1
    MinorVersion        : 0
    Name                : Test1.docx
    ServerRelativeUrl   : /sites/project-x/documents/Test1.docx
    TimeCreated         : 2022-10-30T10:16:18Z
    TimeLastModified    : 2022-10-30T10:16:18Z
    Title               : null
    UIVersion           : 512
    UIVersionLabel      : 1.0
    UniqueId            : b2307a39-e878-458b-bc90-03bc578531d6
    RoleAssignments     : [{"Member":{"Id":3,"IsHiddenInUI":false,"LoginName":"Communication site Owners","Title":"Communication site Owners","PrincipalType":8,"AllowMembersEditMembership":false,"AllowRequestToJoinLeave":false,"AutoAcceptRequestToJoinLeave":false,"Description":null,"OnlyAllowMembersViewMembership":false,"OwnerTitle":"Communication site Owners","RequestToJoinLeaveEmailSetting":""},"RoleDefinitionBindings":[{"BasePermissions":{"High":"2147483647","Low":"4294967295"},"Description":"Has full control.","Hidden":false,"Id":1073741829,"Name":"Full Control","Order":1,"RoleTypeKind":5,"BasePermissionsValue":["ViewListItems","AddListItems","EditListItems","DeleteListItems","ApproveItems","OpenItems","ViewVersions","DeleteVersions","CancelCheckout","ManagePersonalViews","ManageLists","ViewFormPages","AnonymousSearchAccessList","Open","ViewPages","AddAndCustomizePages","ApplyThemeAndBorder","ApplyStyleSheets","ViewUsageData","CreateSSCSite","ManageSubwebs","CreateGroups","ManagePermissions","BrowseDirectories","BrowseUserInfo","AddDelPrivateWebParts","UpdatePersonalWebParts","ManageWeb","AnonymousSearchAccessWebLists","UseClientIntegration","UseRemoteAPIs","ManageAlerts","CreateAlerts","EditMyUserInfo","EnumeratePermissions"],"RoleTypeKindValue":"Administrator"}],"PrincipalId":3}]
    ```

=== "CSV"

    ```csv
    CheckInComment,CheckOutType,ContentTag,CustomizedPageStatus,ETag,Exists,IrmEnabled,Length,Level,LinkingUri,LinkingUrl,MajorVersion,MinorVersion,Name,ServerRelativeUrl,TimeCreated,TimeLastModified,Title,UIVersion,UIVersionLabel,UniqueId,RoleAssignments
    ,2,"{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1,1",0,"""{03E171B6-9DF2-46D9-A7A1-D5FA0BD23E4B},1""",1,,5987,1,,,1,0,Test1.docx,/sites/project-x/documents/Test1.docx,2022-10-30T10:16:18Z,2022-10-30T10:16:18Z,,512,1.0,b2307a39-e878-458b-bc90-03bc578531d6,"[{""Member"":{""Id"":3,""IsHiddenInUI"":false,""LoginName"":""Communication site Owners"",""Title"":""Communication site Owners"",""PrincipalType"":8,""AllowMembersEditMembership"":false,""AllowRequestToJoinLeave"":false,""AutoAcceptRequestToJoinLeave"":false,""Description"":null,""OnlyAllowMembersViewMembership"":false,""OwnerTitle"":""Communication site Owners"",""RequestToJoinLeaveEmailSetting"":""""},""RoleDefinitionBindings"":[{""BasePermissions"":{""High"":""2147483647"",""Low"":""4294967295""},""Description"":""Has full control."",""Hidden"":false,""Id"":1073741829,""Name"":""Full Control"",""Order"":1,""RoleTypeKind"":5,""BasePermissionsValue"":[""ViewListItems"",""AddListItems"",""EditListItems"",""DeleteListItems"",""ApproveItems"",""OpenItems"",""ViewVersions"",""DeleteVersions"",""CancelCheckout"",""ManagePersonalViews"",""ManageLists"",""ViewFormPages"",""AnonymousSearchAccessList"",""Open"",""ViewPages"",""AddAndCustomizePages"",""ApplyThemeAndBorder"",""ApplyStyleSheets"",""ViewUsageData"",""CreateSSCSite"",""ManageSubwebs"",""CreateGroups"",""ManagePermissions"",""BrowseDirectories"",""BrowseUserInfo"",""AddDelPrivateWebParts"",""UpdatePersonalWebParts"",""ManageWeb"",""AnonymousSearchAccessWebLists"",""UseClientIntegration"",""UseRemoteAPIs"",""ManageAlerts"",""CreateAlerts"",""EditMyUserInfo"",""EnumeratePermissions""],""RoleTypeKindValue"":""Administrator""}],""PrincipalId"":3}]"
    ```
