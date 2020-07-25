# spo file move

Moves a file to another location

## Usage

```sh
m365 spo file move [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located

`-s, --sourceUrl <sourceUrl>`
: Site-relative URL of the file to move

`-t, --targetUrl <targetUrl>`
: Server-relative URL where to move the file

`--deleteIfAlreadyExists`
: If a file already exists at the targetUrl, it will be moved to the recycle bin. If omitted, the move operation will be canceled if the file already exists at the targetUrl location

`--allowSchemaMismatch`
: Ignores any missing fields in the target document library and moves the file anyway

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

When you move a file using the `spo file move` command, all of the versions are being moved.

## Examples

Move file to a document library in another site collection

```sh
m365 spo file move --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/
```

Move file to a document library in the same site collection

```sh
m365 spo file move --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test1/HRDocuments/
```

Move file to a document library in another site collection. If a file with the same name already exists in the target document library, move it to the recycle bin

```sh
m365 spo file move --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/ --deleteIfAlreadyExists
```

Move file to a document library in another site collection. Allow for schema mismatch

 ```sh
m365 spo file move --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/ --allowSchemaMismatch
```


## More information

- Move items from a SharePoint document library: [https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc](https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc)