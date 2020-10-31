# spo folder move

Moves a folder to another location

## Usage

```sh
m365 spo folder move [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the folder is located

`-s, --sourceUrl <sourceUrl>`
: Site-relative URL of the folder to move

`-t, --targetUrl <targetUrl>`
: Server-relative URL where to move the folder

`--allowSchemaMismatch`
: Ignores any missing fields in the target destination and moves the folder anyway

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

When you move a folder using the `spo folder move` command, all of the document versions are moved.

## Examples

Move folder to a document library in another site collection

```sh
m365 spo folder move --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test2/Shared%20Documents/
```

Move folder to a document library in the same site collection

```sh
m365 spo folder move --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test1/HRDocuments/
```

Move folder to a document library in another site collection. Allow for schema mismatch

```sh
m365 spo file move --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/ --allowSchemaMismatch
```

## More information

- Move items from a SharePoint document library: [https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc](https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc)