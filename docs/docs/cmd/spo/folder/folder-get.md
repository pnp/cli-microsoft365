# spo folder get

Gets information about the specified folder

## Usage

```sh
m365 spo folder get [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder is located

`-f, --folderUrl <folderUrl>`
: Site-relative URL of the folder

--8<-- "docs/cmd/_global.md"

## Remarks

If no folder exists at the specified URL, you will get a `Please check the folder URL. Folder might not exist on the specified URL` error.

## Examples

Get folder properties for folder with site-relative url _'/Shared Documents'_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder get --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents'
```
