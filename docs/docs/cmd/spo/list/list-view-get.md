# spo list view get

Gets information about specific list view

## Usage

```sh
m365 spo list view get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`--listId [listId]`
: ID of the list where the view is located. Specify only one of `listTitle`, `listId` or `listUrl`

`--listTitle [listTitle]`
: Title of the list where the view is located. Specify only one of `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or web-relative URL of the list where the view is located. Specify only one of `listTitle`, `listId` or `listUrl`

`--id [id]`
: ID of the view to get. Specify `title` or `id` but not both

`--title [title]`
: Title of the view to get. Specify `title` or `id` but not both

--8<-- "docs/cmd/_global.md"

## Examples

Gets a list view by name from a list located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view get --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --title 'All Items'
```

Gets a list view by ID from a list located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view get --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'Lists/My List' --id 330f29c5-5c4c-465f-9f4b-7903020ae1ce
```

Gets a list view by name from a list located in site _https://contoso.sharepoint.com/sites/project-x_. Retrieve the list by its ID

```sh
m365 spo list view get --webUrl https://contoso.sharepoint.com/sites/project-x --listId 330f29c5-5c4c-465f-9f4b-7903020ae1c1 --title 'All Items'
```
