# spo listitem set

Creates a list item in the specified list

## Usage

```sh
spo listitem set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the item should be updated
`-l, --listId [listId]`|ID of the list where the item should be updated. Specify `listId` or `listTitle` but not both
`-t, --listTitle [listTitle]`|Title of the list where the item should be updated. Specify `listId` or `listTitle` but not both
`-i, --id [listItemId]`|ID of the list item to be updated
`-c, --contentType [contentType]`|The name or the ID of the content type to associate with the new item
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To update an item in a list, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Update an item with Id 147 with Title _Demo Item_ and content type name _Item_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem set --contentType Item --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Item"
```

Update an item with Id 147 with Title _Demo Multi Managed Metadata Field_ and a single-select metadata field named _SingleMetadataField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Managed Metadata Field" --SingleMetadataField "TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377;"
```

Update an item with Id 147 with Title _Demo Multi Managed Metadata Field_ and a multi-select metadata field named _MultiMetadataField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Managed Metadata Field" --MultiMetadataField "TermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75;"
```

Update an item with Id 147 with Title _Demo Single Person Field_ and a single-select people field named _SinglePeopleField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Person Field" --SinglePeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'}]"
```

Update an item with Id 147 with Title _Demo Multi Person Field_ and a multi-select people field named _MultiPeopleField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Person Field" --MultiPeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'},{'Key':'i:0#.f|membership|adamb@conotoso.com'}]"
```

Update an item with Id 147 with Title _Demo Hyperlink Field_ and a hyperlink field named _CustomHyperlink_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Hyperlink Field" --CustomHyperlink "https://www.bing.com, Bing"
```