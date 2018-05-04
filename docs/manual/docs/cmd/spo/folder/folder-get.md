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
`-u, --webUrl <webUrl>`|The URL of the site where the folder is
`-f, --folderUrl <folderUrl>`|Site-relative URL of the folder
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

  To get information about a folder, you have to first connect to SharePoint using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Get folder properties for folder with site relative url _'/Shared Documents'_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo folder get --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents'
```

## More information

- Working with folders and files with REST: [https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest)