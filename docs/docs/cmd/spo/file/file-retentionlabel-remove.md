# spo file retentionlabel remove

Clears the retention label from a file

## Usage

```sh
m365 spo file retentionlabel remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the retentionlabel from a file to remove is located

`--fileUrl [fileUrl]`
: The server-relative URL of the file of which the label should be removed. Specify either `fileUrl` or `fileId` but not both.

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file of which the label should be removed. Specify either `fileUrl` or `fileId` but not both.

`--confirm`
: Don't prompt for confirming removing the retention label from a file

--8<-- "docs/cmd/_global.md"

## Examples

Removes the retention label from a file in a given site based on the file id

```sh
m365 spo file retentionlabel remove --webUrl https://contoso.sharepoint.com/sites/project-x --fileId 0cd891ef-afce-4e55-b836-fce03286cccf
```

Removes the retention label from a file in a given site based on the file url

```sh
m365 spo file retentionlabel remove --webUrl https://contoso.sharepoint.com/sites/project-x --fileUrl /sites/project-x/Shared Documents/Document.docx --id 1
```

## Response

The command won't return a response on success.
