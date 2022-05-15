# spo listitem roleinheritance reset

Restores the role inheritance of list item, file, or folder

## Usage

```sh
m365 spo listitem roleinheritance reset [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the item for which to reset role inheritance is located

`--listItemId <listItemId>`
: ID of the item for which to reset role inheritance

`--listId [listId]`
: ID of the list. Specify listId or listTitle but not both

`--listTitle [listTitle]`
: Title of the list. Specify listId or listTitle but not both

--8<-- "docs/cmd/_global.md"

## Examples

Restore role inheritance of list item with id 8 from list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem roleinheritance reset --webUrl https://contoso.sharepoint.com/sites/project-x --listItemId 8 --listId 0cd891ef-afce-4e55-b836-fce03286cccf
```

Restore role inheritance of list item with id 8 from list with title _test_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem roleinheritance reset --webUrl https://contoso.sharepoint.com/sites/project-x --listItemId 8 --listTitle test
```
