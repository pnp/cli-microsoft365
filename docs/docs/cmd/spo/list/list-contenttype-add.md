# spo list contenttype add

Adds content type to list

## Usage

```sh
m365 spo list contenttype add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located.

`-i, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-i, --id <id>`
: ID of the content type to add to the list

--8<-- "docs/cmd/_global.md"

## Examples

Adds a specific existing content type to a list retrieved by id in a specific site.

```sh
m365 spo list contenttype add --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 0x0120
```

Adds a specific existing content type to a list retrieved by title in a specific site.

```sh
m365 spo list contenttype add --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --id 0x0120
```

Adds a specific existing content type to a list retrieved by server relative URL in a specific site.

```sh
m365 spo list contenttype add --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'sites/project-x/Documents' --contentTypeId 0x0120
```
