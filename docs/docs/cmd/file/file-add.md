# file add

Uploads file to the specified site using Microsoft Graph

## Usage

```sh
m365 file add [options]
```

## Options

`-u, --folderUrl <folderUrl>`
: The URL of the document library where the file should be uploaded to

`-p, --filePath <filePath>`
: Local path to the file to upload

--8<-- "docs/cmd/_global.md"

## Remarks

The `folderUrl` must be an absolute URL to the document library where the file should be uploaded. The document library can be located in any site collection in your tenant, including OneDrive for Business. The `folderUrl` can also point to a (sub)folder in the document library.

## Examples

Uploads file from the current folder to the root folder of a document library in the root site collection

```sh
m365 file add --filePath file.pdf --folderUrl "https://contoso.sharepoint.com/Shared Documents"
```

Uploads file from the current folder to a subfolder of a document library in the root site collection

```sh
m365 file add --filePath file.pdf --folderUrl "https://contoso.sharepoint.com/Shared Documents/Folder"
```

Uploads file from the current folder to a document library in OneDrive for Business

```sh
m365 file add --filePath file.pdf --folderUrl "https://contoso-my.sharepoint.com/personal/steve_contoso_com/Documents"
```

Uploads file from the current folder to a document library in a non-root site collection

```sh
m365 file add --filePath file.pdf --folderUrl "https://contoso.sharepoint.com/sites/Contoso/Shared Documents"
```
