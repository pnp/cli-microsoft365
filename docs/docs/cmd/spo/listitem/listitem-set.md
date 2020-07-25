# spo listitem set

Updates a list item in the specified list

## Usage

```sh
m365 spo listitem set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the item should be updated

`-i, --id <id>`
: ID of the list item to update.

`-l, --listId [listId]`
: ID of the list where the item should be updated. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list where the item should be updated. Specify `listId` or `listTitle` but not both

`-c, --contentType [contentType]`
: The name or the ID of the content type to associate with the updated item

`-s, --systemUpdate`
: Update the item without updating the modified date and modified by fields

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Update an item with id _147_ with title _Demo Item_ and content type name _Item_ in list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --contentType Item --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Item"
```

Update an item with id _147_ with title _Demo Multi Managed Metadata Field_ and a single-select metadata field named _SingleMetadataField_ in list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Managed Metadata Field" --SingleMetadataField "TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377;"
```

Update an item with id _147_ with Title _Demo Multi Managed Metadata Field_ and a multi-select metadata field named _MultiMetadataField_ in list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Managed Metadata Field" --MultiMetadataField "TermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75;"
```

Update an item with id 147 with Title _Demo Single Person Field_ and a single-select people field named _SinglePeopleField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Person Field" --SinglePeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'}]"
```

Update an item with id _147_ with Title _Demo Multi Person Field_ and a multi-select people field named _MultiPeopleField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Person Field" --MultiPeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'},{'Key':'i:0#.f|membership|adamb@conotoso.com'}]"
```

Update an item with id _147_ with Title _Demo Hyperlink Field_ and a hyperlink field named _CustomHyperlink_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Hyperlink Field" --CustomHyperlink "https://www.bing.com, Bing"
```