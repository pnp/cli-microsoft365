# spo list contenttype default set

Sets the default content type for a list

## Usage

```sh
m365 spo list contenttype default set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located.

`-l, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-c, --contentTypeId <contentTypeId>`
: ID of the content type

--8<-- "docs/cmd/_global.md"

## Examples

Set a content type with a specific id as default a list retrieved by id located in a specific site.

```sh
m365 spo list contenttype default set --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --contentTypeId 0x0120
```

Set a content type with a specific id as default a list retrieved by title located in a specific site.

```sh
m365 spo list contenttype default set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --contentTypeId 0x0120
```

Set a content type with a specific id as default a list retrieved by server relative URL located in a specific site.

```sh
m365 spo list contenttype default set --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'sites/project-x/Documents' --contentTypeId 0x0120
```

## Response

The command won't return a response on success.
