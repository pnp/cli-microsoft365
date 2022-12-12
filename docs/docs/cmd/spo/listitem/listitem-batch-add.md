# spo listitem batch add

Creates list items in a batch

## Usage

```sh
m365 spo listitem batch add [options]
```

## Options

`-p, --filePath <filePath>`
: The absolute or relative path to a flat file containing the list items.

`-u, --webUrl <webUrl>`
: URL of the site.

`-l, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`, but not multiple.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`, but not multiple.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`, but not multiple.

--8<-- "docs/cmd/_global.md"

## Examples

Add a batch of items to a list retrieved by title in a specific site

```sh
m365 spo listitem batch add --filePath "C:\Path\To\Csv\CsvFile.csv" --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List"
```

Add a batch of items to a list retrieved by Id in a specific site

```sh
m365 spo listitem batch add --filePath "C:\Path\To\Csv\CsvFile.csv" --webUrl https://contoso.sharepoint.com/sites/project-x --listId fe54c47b-22e4-4cab-8a10-3fc54003fb4c
```

Add a batch of items to a list defined by server-relative URL in a specific site

```sh
m365 spo listitem batch add --filePath "C:\Path\To\Csv\CsvFile.csv" --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl "/sites/project-x/lists/Demo List"
```

## Remarks

A sample CSV can be found below. The first line of the CSV-file should contain the internal column names that you wish to set.

```csv
ContentType,Title,SingleChoiceField,MultiChoiceField,SingleMetadataField,MultiMetadataField,SinglePeopleField,MultiPeopleField,CustomHyperlink,NumberField
Item,Title A,Choice 1,Choice 1;#Choice 2,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;Finance|f994a4ac-cf34-448e-a22c-2b35fd9bbffa;,[{'Key':'i:0#.f|membership|markh@contoso.com'}],"[{'Key':'i:0#.f|membership|markh@contoso.com'},{'Key':'i:0#.f|membership|adamb@contoso.com'}]","https://bing.com, URL",5
```

## Response

The command won't return a response on success.
