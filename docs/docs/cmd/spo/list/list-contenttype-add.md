# spo list contenttype add

Adds content type to list

## Usage

```sh
m365 spo list contenttype add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-l, --listId [listId]`
: ID of the list to which to add the content type. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list to which to add the content type. Specify `listId` or `listTitle` but not both

`-c, --contentTypeId <contentTypeId>`
: ID of the content type to add to the list

--8<-- "docs/cmd/_global.md"

## Examples

Add existing content type with ID _0x0120_ to the list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list contenttype add --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --contentTypeId 0x0120
```

Add existing content type with ID _0x0120_ to the list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list contenttype add --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --contentTypeId 0x0120
```
