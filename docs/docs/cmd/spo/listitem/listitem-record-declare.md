# spo listitem record declare

Declares the specified list item as a record

## Usage

```sh
m365 spo listitem record declare [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-l, --listId [listId]`
: The ID of the list where the item is located. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: The title of the list where the item is located. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

`-i, --id <id>`
: The ID of the list item to declare as record

`-d, --date [date]`
: Record declaration date in ISO format. eg. 2019-12-31

--8<-- "docs/cmd/_global.md"

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

Declare a document with a specific id as a record with a record declaration date a list retrieved by server-relative URL located in a specific site

```sh
m365 spo listitem record declare --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl '/sites/project-x/Lists/Demo List' --id 1 --date 2013-09-03
```
