# spo roledefinition list

Gets list of role definitions for the specified site

## Usage

```sh
m365 spo roledefinition list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve role definitions.

--8<-- "docs/cmd/_global.md"

## Examples

Return list of role definitions for the given site

```sh
m365 spo roledefinition list --webUrl https://contoso.sharepoint.com/sites/project-x
```

## Response

=== "JSON"

    ```json
    [
      {
        "BasePermissions": {
          "High": 2147483647,
          "Low": 4294967295
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
    ]
    ```

=== "Text"

    ```text
    Id          Name
    ----------  -----------------------
    1073741829  Full Control
    ```

=== "CSV"

    ```csv
    Id,Name
    1073741829,Full Control
    ```
