# spo listitem get

Gets a list item from the specified list

## Usage

```sh
m365 spo listitem get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site from which the item should be retrieved

`-i, --id <id>`
: ID of the item to retrieve.

`-l, --listId [listId]`
: ID of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: Title of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

`-p, --properties [properties]`
: Comma-separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested

`--withPermissions`
: Set if you want to return associated roles and permissions

--8<-- "docs/cmd/_global.md"

## Remarks

If you want to specify a lookup type in the `properties` option, define which columns from the related list should be returned.

## Examples

Get an item with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get an items _Title_ and _Created_ column with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Created"
```

Get an items _Title_, _Created_ column and lookup column _Company_ with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Created,Company/Title"
```

Get an item with specific properties from a list retrieved by server-relative URL in a specific site

```sh
m365 spo listitem get --listUrl /sites/project-x/documents --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Created,Company/Title"
```

Get an item with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_ with permissions

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --withPermissions
```


## Response

=== "JSON"

    ```json
    {
      "FileSystemObjectType": 0,
      "Id": 147,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ContentTypeId": "0x010078BC2C6F12F0DB41BA554210A2BFA81600A320A11ABE6E90468525ECC747660126",
      "Title": "Demo Item",
      "Modified": "2022-10-30T10:55:37Z",
      "Created": "2022-10-30T10:55:22Z",
      "AuthorId": 10,
      "EditorId": 10,
      "OData__UIVersionString": "3.0",
      "Attachments": false,
      "GUID": "87f3138d-fac3-4126-97c0-543e55672261",
      "ComplianceAssetId": null
    }
    ```

=== "Text"

    ```text
    Attachments             : false
    AuthorId                : 10
    ComplianceAssetId       : null
    ContentTypeId           : 0x010078BC2C6F12F0DB41BA554210A2BFA81600A320A11ABE6E90468525ECC747660126
    Created                 : 2022-10-30T10:55:22Z
    EditorId                : 10
    FileSystemObjectType    : 0
    GUID                    : 87f3138d-fac3-4126-97c0-543e55672261
    Id                      : 147
    Modified                : 2022-10-30T10:55:37Z
    OData__UIVersionString  : 3.0
    ServerRedirectedEmbedUri: null
    ServerRedirectedEmbedUrl:
    Title                   : Demo Item
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,ComplianceAssetId
    0,147,,,0x010078BC2C6F12F0DB41BA554210A2BFA81600A320A11ABE6E90468525ECC747660126,Demo Item,2022-10-30T10:55:37Z,2022-10-30T10:55:22Z,10,10,3.0,,87f3138d-fac3-4126-97c0-543e55672261,
    ```

### `withPermissions` response

When we make use of the option `withPermissions` the response will differ. 

## Response

=== "JSON"

    ```json
    {
      "FileSystemObjectType": 0,
      "Id": 147,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ContentTypeId": "0x010078BC2C6F12F0DB41BA554210A2BFA81600A320A11ABE6E90468525ECC747660126",
      "Title": "Demo Item",
      "Modified": "2022-10-30T10:55:37Z",
      "Created": "2022-10-30T10:55:22Z",
      "AuthorId": 10,
      "EditorId": 10,
      "OData__UIVersionString": "3.0",
      "Attachments": false,
      "GUID": "87f3138d-fac3-4126-97c0-543e55672261",
      "ComplianceAssetId": null,
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
    Attachments             : false
    AuthorId                : 10
    ComplianceAssetId       : null
    ContentTypeId           : 0x010078BC2C6F12F0DB41BA554210A2BFA81600A320A11ABE6E90468525ECC747660126
    Created                 : 2022-10-30T10:55:22Z
    EditorId                : 10
    FileSystemObjectType    : 0
    GUID                    : 87f3138d-fac3-4126-97c0-543e55672261
    Id                      : 147
    Modified                : 2022-10-30T10:55:37Z
    OData__UIVersionString  : 3.0
    RoleAssignments         : [{"Member":{"Id":3,"IsHiddenInUI":false,"LoginName":"Communication site Owners","Title":"Communication site Owners","PrincipalType":8,"AllowMembersEditMembership":false,"AllowRequestToJoinLeave":false,"AutoAcceptRequestToJoinLeave":false,"Description":null,"OnlyAllowMembersViewMembership":false,"OwnerTitle":"Communication site Owners","RequestToJoinLeaveEmailSetting":""},"RoleDefinitionBindings":[{"BasePermissions":{"High":"2147483647","Low":"4294967295"},"Description":"Has full control.","Hidden":false,"Id":1073741829,"Name":"Full Control","Order":1,"RoleTypeKind":5,"BasePermissionsValue":["ViewListItems","AddListItems","EditListItems","DeleteListItems","ApproveItems","OpenItems","ViewVersions","DeleteVersions","CancelCheckout","ManagePersonalViews","ManageLists","ViewFormPages","AnonymousSearchAccessList","Open","ViewPages","AddAndCustomizePages","ApplyThemeAndBorder","ApplyStyleSheets","ViewUsageData","CreateSSCSite","ManageSubwebs","CreateGroups","ManagePermissions","BrowseDirectories","BrowseUserInfo","AddDelPrivateWebParts","UpdatePersonalWebParts","ManageWeb","AnonymousSearchAccessWebLists","UseClientIntegration","UseRemoteAPIs","ManageAlerts","CreateAlerts","EditMyUserInfo","EnumeratePermissions"],"RoleTypeKindValue":"Administrator"}],"PrincipalId":3}]
    ServerRedirectedEmbedUri: null
    ServerRedirectedEmbedUrl:
    Title                   : Demo Item
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,ComplianceAssetId,RoleAssignments
    0,147,,,0x010078BC2C6F12F0DB41BA554210A2BFA81600A320A11ABE6E90468525ECC747660126,Demo Item,2022-10-30T10:55:37Z,2022-10-30T10:55:22Z,10,10,3.0,,87f3138d-fac3-4126-97c0-543e55672261,"[{""Member"":{""Id"":3,""IsHiddenInUI"":false,""LoginName"":""Communication site Owners"",""Title"":""Communication site Owners"",""PrincipalType"":8,""AllowMembersEditMembership"":false,""AllowRequestToJoinLeave"":false,""AutoAcceptRequestToJoinLeave"":false,""Description"":null,""OnlyAllowMembersViewMembership"":false,""OwnerTitle"":""Communication site Owners"",""RequestToJoinLeaveEmailSetting"":""""},""RoleDefinitionBindings"":[{""BasePermissions"":{""High"":""2147483647"",""Low"":""4294967295""},""Description"":""Has full control."",""Hidden"":false,""Id"":1073741829,""Name"":""Full Control"",""Order"":1,""RoleTypeKind"":5,""BasePermissionsValue"":[""ViewListItems"",""AddListItems"",""EditListItems"",""DeleteListItems"",""ApproveItems"",""OpenItems"",""ViewVersions"",""DeleteVersions"",""CancelCheckout"",""ManagePersonalViews"",""ManageLists"",""ViewFormPages"",""AnonymousSearchAccessList"",""Open"",""ViewPages"",""AddAndCustomizePages"",""ApplyThemeAndBorder"",""ApplyStyleSheets"",""ViewUsageData"",""CreateSSCSite"",""ManageSubwebs"",""CreateGroups"",""ManagePermissions"",""BrowseDirectories"",""BrowseUserInfo"",""AddDelPrivateWebParts"",""UpdatePersonalWebParts"",""ManageWeb"",""AnonymousSearchAccessWebLists"",""UseClientIntegration"",""UseRemoteAPIs"",""ManageAlerts"",""CreateAlerts"",""EditMyUserInfo"",""EnumeratePermissions""],""RoleTypeKindValue"":""Administrator""}],""PrincipalId"":3}]"
    ```
