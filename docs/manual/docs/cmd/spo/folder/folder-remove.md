# spo folder remove

Deletes a folder form a site

## Usage

```sh
spo folder remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the folder will be deleted
`-s, --sourceUrl <sourceUrl>`|Site-relative URL of the target folder
`--recycle [recycle]`|Recycles the folder instead of actually deleting it
`--confirm [confirm]`|Don't prompt for confirming removing the folder
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

  To remove a folder, you have to first connect to SharePoint using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

  The remove folder command will remove folder only if it is empty. This is how the SharePoint REST and client.svc APIs work.

## Examples

Removes a folder with site-relative URL _'/Shared Documents/My Folder'_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo folder remove --webUrl https://contoso.sharepoint.com/sites/project-x --sourceUrl '/Shared Documents/My Folder'
```

Moves a folder with site-relative URL _'/Shared Documents/My Folder'_ located in site _https://contoso.sharepoint.com/sites/project-x_ to the recycle bin of the site

```sh
spo folder remove --webUrl https://contoso.sharepoint.com/sites/project-x --sourceUrl '/Shared Documents/My Folder' --recycle
```

## More information

- Working with folders and files with REST: [https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest)