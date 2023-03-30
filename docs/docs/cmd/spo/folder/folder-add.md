# spo folder add

Creates a folder within a parent folder

## Usage

```sh
m365 spo folder add [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder will be created.

`-p, --parentFolderUrl <parentFolderUrl>`
: The server- or site-relative URL of the parent folder.

`-n, --name <name>`
: Name of the new folder to be created

--8<-- "docs/cmd/_global.md"

## Examples

Creates folder in a specific library within the site

```sh
m365 spo folder add --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents' --name 'My Folder Name'
```

Creates folder in a specific folder within the site

```sh
m365 spo folder add --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/sites/project-x/Shared Documents/Reports' --name 'Financial reports'
```
