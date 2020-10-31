# spo listitem remove

Removes the specified list item

## Usage

```sh
m365 spo listitem remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list to remove is located

`-i, --id <id>`
: The ID of the list item to remove.

`-l, --listId [listId]`
: List id of the list to remove. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list to remove. Specify either `listId` or `listTitle` but not both

`--recycle`
: Recycle the list item

`--confirm`
: Don't prompt for confirming removing the list item

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Remove the list item with ID _1_ from list with ID  _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem remove --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf -id 1
```

Remove the list item with ID _1_ from list with title _List 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --id 1
```