# spo list contenttype list

Lists content types configured on the list

## Usage

```sh
m365 spo list contenttype list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site

`-i, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: Relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`.

--8<-- "docs/cmd/_global.md"

## Examples

List all content types configured on a specific list retrieved by id in a specific site.


```sh
m365 spo list contenttype list --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf
```

List all content types configured on a specific list retrieved by title in a specific site.

```sh
m365 spo list contenttype list --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents
```

List all content types configured on a specific list retrieved by server relative URL in a specific site.

```sh
m365 spo list contenttype list --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'sites/project-x/Documents'
```
