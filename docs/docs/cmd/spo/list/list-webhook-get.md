# spo list webhook get

Gets information about the specific webhook

## Usage

```sh
m365 spo list webhook get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-l, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-i, --id [id]`
: ID of the webhook to retrieve

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `id` doesn't refer to an existing webhook, you will get a `404 - "404 FILE NOT FOUND"` error.

## Examples

Return information about a specific webhook which belongs to a list retrieved by ID in a specific site

```sh
m365 spo list webhook get --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id cc27a922-8224-4296-90a5-ebbc54da2e85
```

Return information about a specific webhook which belongs to a list retrieved by Title in a specific site

```sh
m365 spo list webhook get --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --id cc27a922-8224-4296-90a5-ebbc54da2e85
```

Return information about a specific webhook which belongs to a list retrieved by URL in a specific site

```sh
m365 spo list webhook get --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl '/sites/project-x/Documents' --id cc27a922-8224-4296-90a5-ebbc54da2e85
```
