# spo folder remove

Deletes the specified folder

## Usage

```sh
m365 spo folder remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder to be deleted is located

`-f, --folderUrl <folderUrl>`
: Site-relative URL of the folder to delete

`--recycle`
: Recycles the folder instead of actually deleting it

`--confirm`
: Don't prompt for confirming deleting the folder

--8<-- "docs/cmd/_global.md"

## Remarks

The `spo folder remove` command will remove folder only if it is empty. If the folder contains any files, deleting the folder will fail.

## Examples

Removes a folder with site-relative URL _/Shared Documents/My Folder_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder remove --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents/My Folder'
```

Moves a folder with site-relative URL _/Shared Documents/My Folder_ located in site _https://contoso.sharepoint.com/sites/project-x_ to the site recycle bin

```sh
m365 spo folder remove --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents/My Folder' --recycle
```