# spo folder roleinheritance break

Breaks the role inheritance of a folder.

## Usage

```sh
m365 spo folder roleinheritance break [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the folder is located.

`-f, --folderUrl <folderUrl>`
: The site-relative URL or server-relative URL of the folder.

`-c, --clearExistingPermissions`
: Set to clear existing roles from the list item

`--confirm`
: Don't prompt for confirmation to breaking role inheritance of the folder.

--8<-- "docs/cmd/_global.md"

## Examples

Breaks inheritance of folder with site-relative url _Shared Documents/TestFolder_ located in site _https://contoso.sharepoint.com/sites/project-x_ keeping the existing permissions of the folder.

```sh
m365 spo folder roleinheritance reset --webUrl "https://contoso.sharepoint.com/sites/project-x" --folderUrl "Shared Documents/TestFolder"
```

Reset inheritance of folder with server-relative url _/sites/project-x/Shared Documents/TestFolder_ located in site _https://contoso.sharepoint.com/sites/project-x_. It will **not** prompt for confirmation before resetting.

```sh
m365 spo folder roleinheritance reset --webUrl "https://contoso.sharepoint.com/sites/project-x" --folderUrl "/sites/project-x/Shared Documents/TestFolder" --confirm
```
