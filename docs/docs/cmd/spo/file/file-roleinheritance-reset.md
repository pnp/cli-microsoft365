# spo file roleinheritance reset

Restores the role inheritance of a file

## Usage

```sh
m365 spo file roleinheritance reset [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the file is located

`--fileUrl [fileUrl]`
: The server-relative URL of the file to retrieve. Specify either `fileUrl` or `fileId` but not both

`i, --fileId [fileId]`
: The UniqueId (GUID) of the file to retrieve. Specify either `fileUrl` or `fileId` but not both

`--confirm`
: Don't prompt for confirmation to reset role inheritance of the file

--8<-- "docs/cmd/_global.md"

## Examples

Reset inheritance of file with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file roleinheritance reset --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileId "b2307a39-e878-458b-bc90-03bc578531d6"
```

Reset inheritance of file with server-relative url _/sites/project-x/documents/Test1_.docx located in site _https://contoso.sharepoint.com/sites/project-x_. It will **not** prompt for confirmation before resetting.

```sh
m365 spo file roleinheritance reset --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileUrl "/sites/project-x/documents/Test1.docx" --confirm
```
