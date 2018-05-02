# spo folder add

Creates a folder within a parent folder

## Usage

```sh
spo folder add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the folder will be added
`-s, --sourceUrl <sourceUrl>`|Site-relative URL of the parent folder
`-n, --name <name>`|Name of the new folder to be added
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

  To add a folder, you have to first connect to SharePoint using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Adds folder in a parent folder with site relative url _'/Shared Documents'_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo folder add --webUrl https://contoso.sharepoint.com/sites/project-x --sourceUrl '/Shared Documents' --name 'My Folder Name'
```

## More information

- Working with folders and files with REST: [https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest)