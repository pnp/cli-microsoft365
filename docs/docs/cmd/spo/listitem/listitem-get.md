# spo listitem add

Gets a list item from the specified list

## Usage

```sh
m365 spo listitem get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site from which the item should be retrieved

`-i, --id <id>`
: ID of the item to retrieve.

`-l, --listId [listId]`
: ID of the list from which to retrieve the item. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list from which to retrieve the item. Specify `listId` or `listTitle` but not both

`-f, --fields [fields]`
: Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get an item with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x
```


Get an items Title and Created column and with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --fields "Title,Created"
```
