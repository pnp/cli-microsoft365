# spo list roleinheritance reset

Restores role inheritance on list or library

## Usage

```sh
m365 spo list roleinheritance reset [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-i, --listId [listId]`
: ID of the list. Specify either id or title but not both

`-t, --listTitle [listTitle]`
: Title of the list. Specify either id or title but not both

--8<-- "docs/cmd/_global.md"

## Examples

Restore role inheritance of list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list roleinheritance reset --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf
```

Restore role inheritance of list with title _test_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list roleinheritance reset --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle test
```
