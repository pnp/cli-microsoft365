# spo folder copy

Copies a folder to another location

## Usage

```sh
spo folder copy [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-w, --webUrl <webUrl>`|The URL of the site where the folder is located
`-u, --sourceUrl <sourceUrl>`|Site-relative URL of the folder to copy
`-t, --targetUrl <targetUrl>`|Server-relative URL where to copy the folder
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To copy a folder, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

When you use Copy to with documents that have version history, only the latest version is copied.

## Examples

Performs folder copy between two site collections for folder with name _MyFolder_ located in site document library _https://contoso.sharepoint.com/sites/test1/Shared%20Documents_

```sh
spo folder copy --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test2/Shared%20Documents/
```

Performs folder copy between two document libraries in the same site collection for folder with name _MyFolder_ located in site document library _https://contoso.sharepoint.com/sites/test1/Shared%20Documents_

```sh
spo folder copy --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test1/HRDocuments/
```

## More information

- Copy items from a SharePoint document library: [https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc](https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc)

