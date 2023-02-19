# spo list view field set

Updates existing column in an existing view (eg. move to a specific position)

## Usage

```sh
m365 spo list view field set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`--listId [listId]`
: ID of the list where the view is located. Specify either `listId`, `listTitle`, or `listUrl`.

`--listTitle [listTitle]`
: Title of the list where the view is located. Specify either `listId`, `listTitle`, or `listUrl`.

 `--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listId` , `listTitle` or `listUrl`.

`--viewId [viewId]`
: ID of the view to update. Specify `viewTitle` or `viewId` but not both

`--viewTitle [viewTitle]`
: Title of the view to update. Specify `viewTitle` or `viewId` but not both

`--id [id]`
: ID of the field to update. Specify `id` or `title` but not both

`--title [title]`
: The **case-sensitive** internal name or display name of the field to update. Specify `id` or `title` but not both

`--position <position>`
: The zero-based index of the position to which to move the field

--8<-- "docs/cmd/_global.md"

## Examples

Move field with ID _330f29c5-5c4c-465f-9f4b-7903020ae1ce_ to the front in view with ID _3d760127-982c-405e-9c93-e1f76e1a1110_ of list with ID _1f187321-f086-4d3d-8523-517e94cc9df9_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view field set --webUrl https://contoso.sharepoint.com/sites/project-x --listId 1f187321-f086-4d3d-8523-517e94cc9df9 --viewId 3d760127-982c-405e-9c93-e1f76e1a1110 --id 330f29c5-5c4c-465f-9f4b-7903020ae1ce --position 0
```

Move field with title _Custom field_ to position _1_ in view with title _All Documents_ of the list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view field set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --viewTitle 'All Documents' --title 'Custom field' --position 1
```

Move field with title _Custom field_ to position _1_ in view with title _All Documents_ of the list with url _/sites/project-x/lists/Events_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view field set --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl '/sites/project-x/lists/Events' --viewTitle 'All Documents' --fieldTitle 'Custom field' --fieldPosition 1
```

Move field with title _Custom field_ to position _1_ in view with title _All Documents_ of the list with site-relative URL _/Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view field set --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl '/Shared Documents' --viewTitle 'All Documents' --fieldTitle 'Custom field' --fieldPosition 1
```

## Response

The command won't return a response on success.
