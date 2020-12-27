# spo folder list

Returns all folders under the specified parent folder

## Usage

```sh
m365 spo folder list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folders to list are located

`-p, --parentFolderUrl <parentFolderUrl>`
: Site-relative URL of the parent folder

--8<-- "docs/cmd/_global.md"

## Examples

Gets list of folders under a parent folder with site-relative url _/Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents'
```