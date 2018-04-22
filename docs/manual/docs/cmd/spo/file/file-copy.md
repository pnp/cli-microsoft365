# spo file copy

Copies a file to another location

## Usage

```sh
spo file copy [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-w, --webUrl <webUrl>`|The URL of the site where the file is located
`-u, --sourceUrl <sourceUrl>`|Site-relative URL of the file to copy
`-t, --targetUrl <targetUrl>`|Server-relative URL where to copy the file
`-d, --deleteIfAlreadyExists [deleteIfAlreadyExists]`|If a file already exists at the targetUrl, it will be moved to the recycle bin. If ommitted, the copy operation will be canceled if the file already exists at the targetUrl location
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To copy a file, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

When you use Copy to with documents that have version history, only the latest version is copied.

## Examples

Performs file copy between two site collections for file with name _sp1.pdf_ located in site document library _https://contoso.sharepoint.com/sites/test1/Shared%20Documents_

```sh
spo file copy --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/
```

Performs file copy between two document libraries in the same site collection for file with name _sp1.pdf_ located in site document library _https://contoso.sharepoint.com/sites/test1/Shared%20Documents_

```sh
spo file copy --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test1/HRDocuments/
```

Performs file copy between two site collections for file with name _sp1.pdf_ located in site document library _https://contoso.sharepoint.com/sites/test1/Shared%20Documents_ with option --deleteIfAlreadyExists. This will delete existing file with the same name in the target folder.

```sh
spo file copy --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/ --deleteIfAlreadyExists
```

## More information

- Copy items from a SharePoint document library: [https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc](https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc)

