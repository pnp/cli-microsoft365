# spo listitem batch set

Updates list items in a batch

## Usage

```sh
m365 spo listitem batch set [options]
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

`--idColumn [idColumn]`
: Name of the column in the csv containing the IDs of the items to set. Defaults to `ID`.

`-s, --systemUpdate`
: Update the item without updating the modified date and modified by fields.

--8<-- "docs/cmd/_global.md"

## Examples

Updates a batch of items to a list retrieved by title in a specific site

```sh
m365 spo listitem batch set --filePath "C:\Path\To\Csv\CsvFile.csv" --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List"
```

Updates a batch of items to a list retrieved by Id in a specific site with a specific `idColumn`

```sh
m365 spo listitem batch set --filePath "C:\Path\To\Csv\CsvFile.csv" --webUrl https://contoso.sharepoint.com/sites/project-x --listId fe54c47b-22e4-4cab-8a10-3fc54003fb4c --idColumn id
```

Updates a batch of items to a list defined by server-relative URL in a specific site without updating the modified details

```sh
m365 spo listitem batch set --filePath "C:\Path\To\Csv\CsvFile.csv" --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl "/sites/project-x/lists/Demo List" --systemUpdate
```

## Remarks

A sample CSV can be found below. The first line of the CSV-file should contain the internal column names that you wish to set.

```csv
ID,ContentType,Title,SingleChoiceField,MultiChoiceField,SingleMetadataField,MultiMetadataField,SinglePeopleField,MultiPeopleField,CustomHyperlink,NumberField,LookupList,LookupListMulti
5,Item,Title Update,Choice 1,Choice 1;#Choice 2,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;Finance|f994a4ac-cf34-448e-a22c-2b35fd9bbffa,john@contoso.com,john@contoso.com;doe@contoso.com,"https://bing.com, URL",5,2,2;3
```

## Response

The command won't return a response on success.
