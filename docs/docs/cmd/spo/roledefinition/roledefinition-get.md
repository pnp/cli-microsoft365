# spo roledefinition get

Gets specified role definition from web

## Usage

```sh
m365 spo roledefinition get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve the role definition.

`-i, --id <id>`
: The Id of the role definition to retrieve.

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve the role definition for the given site

```sh
m365 spo roledefinition get --webUrl https://contoso.sharepoint.com/sites/project-x --id 1
```

## Response

=== "JSON"

    ```json
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
    ```

=== "Text"

    ```text
    BasePermissions     : {"High":"2147483647","Low":"4294967295"}
    BasePermissionsValue: ["ViewListItems","AddListItems","EditListItems","DeleteListItems","ApproveItems","OpenItems","ViewVersions","DeleteVersions","CancelCheckout","ManagePersonalViews","ManageLists","ViewFormPages","AnonymousSearchAccessList","Open","ViewPages","AddAndCustomizePages","ApplyThemeAndBorder","ApplyStyleSheets","ViewUsageData","CreateSSCSite","ManageSubwebs","CreateGroups","ManagePermissions","BrowseDirectories","BrowseUserInfo","AddDelPrivateWebParts","UpdatePersonalWebParts","ManageWeb","AnonymousSearchAccessWebLists","UseClientIntegration","UseRemoteAPIs","ManageAlerts","CreateAlerts","EditMyUserInfo","EnumeratePermissions"]
    Description         : Has full control.
    Hidden              : false
    Id                  : 1073741829
    Name                : Full Control
    Order               : 1
    RoleTypeKind        : 5
    RoleTypeKindValue   : Administrator
    ```

=== "CSV"

    ```csv
    BasePermissions,Description,Hidden,Id,Name,Order,RoleTypeKind,BasePermissionsValue,RoleTypeKindValue
    "{""High"":""2147483647"",""Low"":""4294967295""}",Has full control.,,1073741829,Full Control,1,5,"[""ViewListItems"",""AddListItems"",""EditListItems"",""DeleteListItems"",""ApproveItems"",""OpenItems"",""ViewVersions"",""DeleteVersions"",""CancelCheckout"",""ManagePersonalViews"",""ManageLists"",""ViewFormPages"",""AnonymousSearchAccessList"",""Open"",""ViewPages"",""AddAndCustomizePages"",""ApplyThemeAndBorder"",""ApplyStyleSheets"",""ViewUsageData"",""CreateSSCSite"",""ManageSubwebs"",""CreateGroups"",""ManagePermissions"",""BrowseDirectories"",""BrowseUserInfo"",""AddDelPrivateWebParts"",""UpdatePersonalWebParts"",""ManageWeb"",""AnonymousSearchAccessWebLists"",""UseClientIntegration"",""UseRemoteAPIs"",""ManageAlerts"",""CreateAlerts"",""EditMyUserInfo"",""EnumeratePermissions""]",Administrator
    ```
