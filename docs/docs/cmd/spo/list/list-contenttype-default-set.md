# spo list contenttype default set

Sets the default content type for a list

## Usage

```sh
m365 spo list contenttype default set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-l, --listId [listId]`
: ID of the list on which to set default content type, specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list on which to set default content type, specify `listId` or `listTitle` but not both

`-c, --contentTypeId <contentTypeId>`
: ID of the content type to set as default on the list

--8<-- "docs/cmd/_global.md"

## Examples

Set content type with ID _0x0120_ as default in the list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list contenttype default set --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --contentTypeId 0x0120
```

Set content type with ID _0x0120_ as default in the list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list contenttype default set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --contentTypeId 0x0120
```
