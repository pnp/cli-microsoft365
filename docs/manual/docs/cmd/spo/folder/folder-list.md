# spo folder list

Returns all folders under parent folder

## Usage

```sh
spo folder list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the folders are
`-f, --folderUrl <folderUrl>`|Site-relative URL of the parent folder
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

  To get list of folders under parent folder, you have to first connect to SharePoint using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Gets list of folders under parent folder with site relative url _'/Shared Documents'_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents'
```

## More information

- Working with folders and files with REST: [https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest)