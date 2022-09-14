# spo listitem roleinheritance break

Break inheritance of list item.

## Usage

```sh
m365 spo listitem roleinheritance break [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the item for which to break role inheritance is located

`--listItemId <listItemId>`
: ID of the item for which to break role inheritance

`-l, --listId [listId]`
: ID of the list. Specify listId or listTitle but not both

`-t, --listTitle [listTitle]`
: Title of the list. Specify listId or listTitle but not both

`-c, --clearExistingPermissions`
: Set to clear existing roles from the list item

--8<-- "docs/cmd/_global.md"

## Remarks

By default, when breaking permissions inheritance, the list item will retain existing permissions. To remove existing permissions, use the `--clearExistingPermissions` option.

## Examples

Break inheritance of list item _1_ in list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "_someList_" --listItemId 1
```

Break inheritance of list item _1_ in list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x --listId 202b8199-b9de-43fd-9737-7f213f51c991 --listItemId 1
```

Break inheritance of list item _1_ in list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_ with clearing permissions

```sh
m365 spo listitem roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "_someList_" --listItemId 1 --clearExistingPermissions
```

Break inheritance of list item _1_ in list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ located in site _https://contoso.sharepoint.com/sites/project-x_ with clearing permissions

```sh
m365 spo listitem roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x --listId 202b8199-b9de-43fd-9737-7f213f51c991 --listItemId 1 --clearExistingPermissions
```
