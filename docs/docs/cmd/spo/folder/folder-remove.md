# spo folder remove

Deletes the specified folder

## Usage

```sh
m365 spo folder remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder to be deleted is located.

`-f, --url <url>`
: The server- or site-relative URL of the folder to delete.

`--recycle`
: Recycles the folder instead of actually deleting it.

`--confirm`
: Don't prompt for confirming deleting the folder.

--8<-- "docs/cmd/_global.md"

## Remarks

The `spo folder remove` command will remove folder only if it is empty. If the folder contains any files, deleting the folder will fail.

## Examples

Remove a folder with a specific site-relative URL

```sh
m365 spo folder remove --webUrl https://contoso.sharepoint.com/sites/project-x --url '/Shared Documents/My Folder'
```

Remove a folder with a specific server relative URL to the site recycle bin

```sh
m365 spo folder remove --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/Shared Documents/My Folder' --recycle
```
