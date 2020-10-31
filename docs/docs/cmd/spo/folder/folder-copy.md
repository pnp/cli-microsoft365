# spo folder copy

Copies a folder to another location

## Usage

```sh
m365 spo folder copy [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the folder is located

`-s, --sourceUrl <sourceUrl>`
: Site-relative URL of the folder to copy

`-t, --targetUrl <targetUrl>`
: Server-relative URL where to copy the folder

`--allowSchemaMismatch`
: Ignores any missing fields in the target document library and copies the folder anyway

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

When you copy a folder with documents that have version history, only the latest document version is copied.

## Examples

Copies folder from a document library located in one site collection to another site collection

```sh
m365 spo folder copy --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test2/Shared%20Documents/
```

Copies folder from a document library to another site in the same site collection

```sh
m365 spo folder copy --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test1/HRDocuments/
```

Copies folder to a document library in another site collection. Allow for schema mismatch

```sh
m365 spo folder cope --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test2/Shared%20Documents/ --allowSchemaMismatch
```

## More information

- Copy items from a SharePoint document library: [https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc](https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc)