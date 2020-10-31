# spo listitem record declare

Declares the specified list item as a record

## Usage

```sh
m365 spo listitem record declare [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-l, --listId [listId]`
: The ID of the list where the item is located. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: The title of the list where the item is located. Specify `listId` or `listTitle` but not both

`-i, --id <id>`
: The ID of the list item to declare as record

`-d, --date [date]`
: Record declaration date in ISO format. eg. 2019-12-31

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Declare a document with id _1_ as a record in list with title _Demo List_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem record declare --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List" --id 1
```

Declare a document with id _1_ as a record in list with id _ea8e1109-2013-1a69-bc05-1403201257fc_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem record declare --webUrl https://contoso.sharepoint.com/sites/project-x --listId ea8e1109-2013-1a69-bc05-1403201257fc --id 1
```

Declare a document with id _1_ as a record with record declaration date _March 14, 2012_ in list with title _Demo List_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem record declare --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List" --id 1 --date 2012-03-14
```

Declare a document with id _1_ as a record with record declaration date _September 3, 2013_ in list with id _ea8e1356-5910-abc9-bc05-2408198057fc_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem record declare --webUrl https://contoso.sharepoint.com/sites/project-x --listId ea8e1356-5910-abc9-bc05-2408198057fc --id 1 --date 2013-09-03
```