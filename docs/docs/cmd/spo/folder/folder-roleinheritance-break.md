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
: Clear all existing permissions from the folder.

`--confirm`
: Don't prompt for confirmation to breaking role inheritance of the folder.

--8<-- "docs/cmd/_global.md"

## Examples

Break the inheritance of a folder with a specified site-relative URL.

```sh
m365 spo folder roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --folderUrl "Shared Documents/TestFolder"
```

Break the inheritance of a folder with a specified server-relative URL. It will clear the existing permissions of the folder. It will **not** prompt for confirmation before breaking the inheritance.

```sh
m365 spo folder roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --folderUrl "/sites/project-x/Shared Documents/TestFolder" --clearExistingPermissions --confirm
```
