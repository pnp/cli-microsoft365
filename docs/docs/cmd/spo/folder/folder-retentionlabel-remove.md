# spo folder retentionlabel remove

Clears the retention label from a folder

## Usage

```sh
m365 spo folder retentionlabel remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: The url of the web.

`--folderUrl [folderUrl]`
: The server-relative URL of the folder of which the label should be removed. Specify either `folderUrl` or `folderId` but not both.

`-i, --folderId [folderId]`
: The UniqueId (GUID) of the folder of which the label should be removed. Specify either `folderUrl` or `folderId` but not both.

`--confirm`
: Don't prompt for confirming to remove the label.

--8<-- "docs/cmd/_global.md"

## Remarks

Removing a retentionlabel is only supported on subfolders. Removing a retentionlabel from a rootfolder is currently not supported by by this command:

## Examples

Removes the retention label from a folder in a given site based on the folder id

```sh
m365 spo folder retentionlabel remove --webUrl https://contoso.sharepoint.com/sites/project-x --folderId 0cd891ef-afce-4e55-b836-fce03286cccf
```

Removes the retention label from a folder in a given site based on the folder url

```sh
m365 spo folder retentionlabel remove --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl /sites/project-x/Shared Documents/Folder --id 1
```

## Response

The command won't return a response on success.