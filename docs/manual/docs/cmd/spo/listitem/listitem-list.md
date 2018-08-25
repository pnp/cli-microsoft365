# spo listitem list

Gets a list of items from the specified list

## Usage

```sh
spo listitem list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site from which the item should be retrieved
`-i, --id <id>`|ID of the list to retrieve items from. Specify `id` or `title` but not both
`-t, --title [listTitle]`|Title of the list from which to retrieve the item. Specify `id` or `title` but not both
`-q, --query [camlQuery]`|CAML query to use to query the list of items with
`-f, --fields [fields]`|Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested. Specify `query` or `fields` but not both  
`-l, --filter [odataFilter]`|ODATA filter to use to query the list of items with. Specify `query` or `filter` but not both
`-p, --pageSize [pageSize]`|Number of list items to return. Specify `query` or `pageSize` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get a list of items from a list, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Get a list of items from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem list --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get a list of items from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_ using the CAML query _<Query><View><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo list item</Value></Eq></Where></View></Query>_

```sh
spo listitem list --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --query "<Query><View><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo list item</Value></Eq></Where></View></Query>"
```

Get a list of items from list with a GUID of _935c13a0-cc53-4103-8b48-c1d0828eaa7f_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem list --id 935c13a0-cc53-4103-8b48-c1d0828eaa7f --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get a list of items from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_, specifying fields _ID,Title,Modified_

```sh
spo listitem list --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --fields "ID,Title,Modified"
```

Get a list of items from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_, with an ODATA filter _Title eq 'Demo list item'_

```sh
spo listitem list --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --filter "Title eq 'Demo list item'"
```

Get a list of items from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_, with a page size of _10_

```sh
spo listitem list --title "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --pageSize 10
```

