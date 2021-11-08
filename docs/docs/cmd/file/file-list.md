# file list

Retrieves files from the specified folder and site

## Usage

```sh
m365 file list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder from which to retrieve files is located

`-f, --folderUrl <folderUrl>`
: The server- or site-relative URL of the folder from which to retrieve files

`--recursive`
: Set to retrieve files from subfolders

--8<-- "docs/cmd/_global.md"

## Remarks

This command is an improved version of the `spo file list` command. The main difference between the two commands is, that `file list` uses Microsoft Graph and properly supports retrieving files from large folders. Because `file list` uses Microsoft Graph and `spo file list` uses SharePoint REST APIs, the data returned by both commands is different.

## Examples

Return all files from folder _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 file list --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents'
```

Return all files from the folder _Shared Documents_ and all the sub-folders of _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 file list --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --recursive
```

Return all files from the _Important_ folder in the _Shared Documents_ document library located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 file list --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents/Important'
```
