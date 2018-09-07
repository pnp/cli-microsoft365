# spo folder list

Returns all folders under the specified parent folder

## Usage

```sh
spo folder list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the folders to list are located
`-p, --parentFolderUrl <parentFolderUrl>`|Site-relative URL of the parent folder
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To get list of folders under parent folder, you have to first log in to SharePoint using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Gets list of folders under a parent folder with site-relative url _/Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents'
```