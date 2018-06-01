# spo folder rename

Renames a folder

## Usage

```sh
spo folder rename [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the folder to be renamed is located
`-f, --folderUrl <folderUrl>`|Site-relative URL of the folder (including the folder)
`-n, --name`|New name for the target folder
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To rename a folder, you have to first connect to SharePoint using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Renames a folder with site-relative URL _/Shared Documents/My Folder 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo folder rename --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents/My Folder 1' --name 'My Folder 2'
```