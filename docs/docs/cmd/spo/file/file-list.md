# spo file list

Gets all files within the specified folder and site

## Usage

```sh
m365 spo file list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder from which to retrieve files is located

`-f, --folder <folder>`
: The server- or site-relative URL of the folder from which to retrieve files

--8<-- "docs/cmd/_global.md"

## Examples

Return all files from folder _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file list --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents'
```