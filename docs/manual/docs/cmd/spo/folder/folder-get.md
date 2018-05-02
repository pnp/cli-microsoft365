# spo folder get

Gets information about the specified folder

## Usage

```sh
spo folder get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the folder is located
`-f, --folderUrl <folderUrl>`|Site-relative URL of the folder
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get information about a folder, you have to first connect to SharePoint using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If no folder exists at the specified URL, you will get a `Please check the folder URL. Folder might not exist on the specified URL` error.

## Examples

Get folder properties for folder with site-relative url _'/Shared Documents'_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo folder get --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents'
```