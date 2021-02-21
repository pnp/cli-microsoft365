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
: ID of the list where the view is located. Specify `listTitle` or `listId` but not both

`--listTitle [listTitle]`
: Title of the list where the view is located. Specify `listTitle` or `listId` but not both

`--viewId [viewId]`
: ID of the view to update. Specify `viewTitle` or `viewId` but not both

`--viewTitle [viewTitle]`
: Title of the view to update. Specify `viewTitle` or `viewId` but not both

`--fieldId [fieldId]`
: ID of the field to update. Specify `fieldId` or `fieldTitle` but not both

`--fieldTitle [fieldTitle]`
: The **case-sensitive** internal name or display name of the field to update. Specify `fieldId` or `fieldTitle` but not both

`--fieldPosition <fieldPosition>`
: The zero-based index of the position to which to move the field

--8<-- "docs/cmd/_global.md"

## Examples

Move field with ID _330f29c5-5c4c-465f-9f4b-7903020ae1ce_ to the front in view with ID _3d760127-982c-405e-9c93-e1f76e1a1110_ of list with ID _1f187321-f086-4d3d-8523-517e94cc9df9_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view field set --webUrl https://contoso.sharepoint.com/sites/project-x --listId 1f187321-f086-4d3d-8523-517e94cc9df9 --viewId 3d760127-982c-405e-9c93-e1f76e1a1110 --fieldId 330f29c5-5c4c-465f-9f4b-7903020ae1ce --fieldPosition 0
```

Move field with title _Custom field_ to position _1_ in view with title _All Documents_ of the list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view field set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --viewTitle 'All Documents' --fieldTitle 'Custom field' --fieldPosition 1
```
