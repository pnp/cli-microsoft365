# spo list view remove

Deletes the specified view from the list

## Usage

```sh
m365 spo list view remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list to remove the view from is located

`--listId [listId]`
: ID of the list from which the view should be removed. Specify either `listId` or `listTitle` but not both

`--listTitle [listTitle]`
: Title of the list from which the view should be removed. Specify either `listId` or `listTitle` but not both

`--viewId [viewId]`
: ID of the view to remove. Specify either `viewId` or `viewTitle` but not both

`--viewTitle [viewTitle]`
: Title of the view to remove. Specify either `viewId` or `viewTitle` but not both

`--confirm`
: Don't prompt for confirming removing the view

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Remove view with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ from the list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view remove --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --viewId cc27a922-8224-4296-90a5-ebbc54da2e81
```

Remove view with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ from the list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --viewId cc27a922-8224-4296-90a5-ebbc54da2e81
```

Remove view with title _MyView_ from a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list view remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --viewTitle MyView
```

Remove view with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ from a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_ without being asked for confirmation

```sh
m365 spo list view remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --viewId cc27a922-8224-4296-90a5-ebbc54da2e81 --confirm
```