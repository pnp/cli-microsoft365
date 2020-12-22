# spo folder rename

Renames a folder

## Usage

```sh
m365 spo folder rename [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder to be renamed is located

`-f, --folderUrl <folderUrl>`
: Site-relative URL of the folder (including the folder)

`-n, --name`
: New name for the target folder

--8<-- "docs/cmd/_global.md"

## Examples

Renames a folder with site-relative URL _/Shared Documents/My Folder 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder rename --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents/My Folder 1' --name 'My Folder 2'
```