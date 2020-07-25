# spo listitem isrecord

Checks if the specified list item is a record

## Usage

```sh
m365 spo listitem isrecord [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the list is located

`-i, --id <id>`
: The ID of the list item to check if it is a record

`-l, --listId [listId]`
: The ID of the list where the item is located. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: The title of the list where the item is located. Specify `listId` or `listTitle` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Check whether the document with id _1_ in list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_ is a record

```sh
m365 spo listitem isrecord --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'Documents' --id 1
```

Check whether the document with id _1_ in list with id _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_ is a record

```sh
m365 spo listitem isrecord --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 1
```
