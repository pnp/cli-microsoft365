# spo list view list

Lists views configured on the specified list

## Usage

```sh
m365 spo list view list [options]
```

## Options

 `-u, --webUrl <webUrl>`
: URL of the site where the list is located

 `-i, --listId [listId]`
: ID of the list for which to list configured views. Specify either `listId`, `listTitle`, or `listUrl`.

 `-t, --listTitle [listTitle]`
: Title of the list for which to list configured views. Specify either `listId`, `listTitle`, or `listUrl`.

 `--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listId` , `listTitle` or `listUrl`.

--8<-- "docs/cmd/_global.md"

## Examples

List all views for a list by title

```sh
m365 spo list view list --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents
```

List all views for a list by ID

```sh
m365 spo list view list --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf
```

List all views for a list by url

```sh
m365 spo list view list --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl '/sites/project-x/lists/Events'
```

