# spo listitem batch add

Creates a list item in the specified list

## Usage

```sh
spo listitem batch add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the item should be added
`-l, --listId [listId]`|ID of the list where the item should be added. Specify `listId` or `listTitle` but not both
`-t, --listTitle [listTitle]`|Title of the list where the item should be added. Specify `listId` or `listTitle` but not both
`-p, --path [path]`|  The path of the csv file with records to be added to the SharePoint list 
`-c, --contentType [contentType]`|The name or the ID of the content type to associate with the new item
`-f, --folder [folder]`|The list-relative URL of the folder where the item should be created
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Usage Notes

The first row of the csv file contains column headers. The column headers must match the internal name of the field
in the SharePoint List.

For a detailed explanation on how to format the fields within the csv, use spo help listitem command add.



## Examples
Add an item with content type name _Item_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_ for each row in test.csv

```sh
spo listitem add --contentType Item --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --path  .\test.csv
```

