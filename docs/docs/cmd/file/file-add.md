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

`--siteUrl [siteUrl]`
: URL of the site to which upload the file. Specify to suppress lookup.

--8<-- "docs/cmd/_global.md"

## Remarks

The `folderUrl` must be an absolute URL to the document library where the file should be uploaded. The document library can be located in any site collection in your tenant, including OneDrive for Business. The `folderUrl` can also point to a (sub)folder in the document library.

By default, the `file add` command will automatically lookup the ID of the site where you want to upload the file based on the specified `folderUrl`. It will do this, by breaking the URL into chunks and incrementally calling Microsoft Graph to retrieve site information. This is necessary, because there is no other way looking at the URL to distinguish where the site URL ends and the document library URL starts. If you want to speed up uploading files, or you use resource-specific consent and your Azure AD app only has access to the specific site, you can use the `siteUrl` option to specify the URL of the site yourself.

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

Uploads file from the current folder to a document library in the specified site

```sh
m365 file add --filePath file.pdf --folderUrl "https://contoso.sharepoint.com/sites/Contoso/Shared Documents" --siteUrl "https://contoso.sharepoint.com/sites/Contoso"
```
