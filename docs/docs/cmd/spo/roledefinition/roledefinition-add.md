# spo roledefinition add

Adds a new roledefinition to web

## Usage

```sh
m365 spo roledefinition add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site to which role should be added

`-n, --name <name>`
: role definition name

`-d, --description [description]`
: role definition description

`--rights [rights]`
: A case-sensitive string array that contain the permissions needed for the custom action. Allowed values `EmptyMask,ViewListItems,AddListItems,EditListItems,DeleteListItems,ApproveItems,OpenItems,ViewVersions,DeleteVersions,CancelCheckout,ManagePersonalViews,ManageLists,ViewFormPages,AnonymousSearchAccessList,Open,ViewPages,AddAndCustomizePages,ApplyThemeAndBorder,ApplyStyleSheets,ViewUsageData,CreateSSCSite,ManageSubwebs,CreateGroups,ManagePermissions,BrowseDirectories,BrowseUserInfo,AddDelPrivateWebParts,UpdatePersonalWebParts,ManageWeb,AnonymousSearchAccessWebLists,UseClientIntegration,UseRemoteAPIs,ManageAlerts,CreateAlerts,EditMyUserInfo,EnumeratePermissions,FullMask`. Default `EmptyMask`

--8<-- "docs/cmd/_global.md"

## Remarks

The `--rights` option accepts **case-sensitive** values.

## Examples

Adds the role definition for site _https://contoso.sharepoint.com/sites/project-x_ with name _test_

```sh
m365 spo roledefinition add --webUrl https://contoso.sharepoint.com/sites/project-x --name test
```

Adds the role definition for site _https://contoso.sharepoint.com/sites/project-x_ with name _test_ and description _test description_ and rights _ViewListItems,AddListItems,EditListItems,DeleteListItems_

```sh
m365 spo roledefinition add --webUrl https://contoso.sharepoint.com/sites/project-x --name test --description "test description" --rights "ViewListItems,AddListItems,EditListItems,DeleteListItems"
```
