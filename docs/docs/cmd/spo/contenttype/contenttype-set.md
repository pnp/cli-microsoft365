# spo contenttype set

Update an existing content type

## Usage

```sh
m365 spo contenttype set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the content type to update is defined.

`-i, --id [id]`
: ID of the content type to update. Specify `id` or `name` but not both, one is required.

`-n, name [name]`
: Name of the content type to update. Specify the `id` or the `name` but not both, one is required.

`--listTitle [listTitle]`
: Title of the list if you want to update a list content type. Specify either `listTitle`, `listId` or `listUrl`.

`--listId [listId]`
: ID of the list if you want to update a list content type. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: URL of the list if you want to update a list content type. Specify either `listTitle`, `listId` or `listUrl`.

`--updateChildren`
: Specify if you want to push updates to child content types.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important "Updating child content types"
    When specifying the `--updateChildren` flag, SharePoint will only propagate the changes that are made in the current request. If you want to know more about updating a content type and propagating changes to child content types, be sure to [read more here](https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ms442695(v=office.14)#considerations-when-updating-child-content-types). 

## Examples

Move site content type to a different group

```sh
m365 spo contenttype set --id 0x001001 --webUrl https://contoso.sharepoint.com --Group "My group"
```

Rename list content type

```sh
m365 spo contenttype set --name "My old item" --webUrl https://contoso.sharepoint.com --listTitle "My list" --Name "My item"
```

Configure a form customizer with manifest ID _19890cce-15d8-4af9-bfcb-72da06d13ed8_ on a site content type and push changes to child content types.

```sh
m365 spo contenttype set --name "My content type" --webUrl https://contoso.sharepoint.com --DisplayFormClientSideComponentId "19890cce-15d8-4af9-bfcb-72da06d13ed8" --EditFormClientSideComponentId "19890cce-15d8-4af9-bfcb-72da06d13ed8" --NewFormClientSideComponentId "19890cce-15d8-4af9-bfcb-72da06d13ed8" --updateChildren
```
