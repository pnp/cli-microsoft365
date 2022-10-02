# spo folder roleinheritance reset

Restores the role inheritance of a folder

## Usage

```sh
m365 spo folder roleinheritance reset [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the folder is located

`-f, --folderUrl <folderUrl>`
: The site-relative URL of the folder

`--confirm`
: Don't prompt for confirmation to reset role inheritance of the folder

--8<-- "docs/cmd/_global.md"

## Examples

Reset inheritance of folder with site-relative url _Shared Documents/TestFolder_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder roleinheritance reset --webUrl "https://contoso.sharepoint.com/sites/project-x" --folderUrl "Shared Documents/TestFolder"
```

Reset inheritance of folder with site-relative url _Shared Documents/TestFolder_ located in site _https://contoso.sharepoint.com/sites/project-x_. It will **not** prompt for confirmation before resetting.

```sh
m365 spo folder roleinheritance reset --webUrl "https://contoso.sharepoint.com/sites/project-x" --folderUrl "Shared Documents/TestFolder" --confirm
```