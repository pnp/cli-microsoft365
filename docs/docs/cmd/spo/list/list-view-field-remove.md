# spo list view field remove

Removes the specified field from list view

## Usage

```sh
m365 spo list view field remove [options]
```

## Options

`-h, --help`
: output usage information

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
: ID of the field to remove. Specify fieldId or fieldTitle but not both

`--fieldTitle [fieldTitle]`
: The **case-sensitive** internal name or display name of the field to remove. Specify `fieldId` or `fieldTitle` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Remove field with ID _330f29c5-5c4c-465f-9f4b-7903020ae1ce_ from view with ID _3d760127-982c-405e-9c93-e1f76e1a1110_ from the list with ID _1f187321-f086-4d3d-8523-517e94cc9df9_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view field remove --webUrl https://contoso.sharepoint.com/sites/project-x --listId 1f187321-f086-4d3d-8523-517e94cc9df9 --viewId 3d760127-982c-405e-9c93-e1f76e1a1110 --fieldId 330f29c5-5c4c-465f-9f4b-7903020ae1ce
```

Remove field with title _Custom field_ from view with title _Custom view_ from the list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view field remove --webUrl https://contoso.sharepoint.com/sites/project-x --fieldTitle 'Custom field' --listTitle Documents --viewTitle 'Custom view'
```