# spo folder get

Gets information about the specified folder

## Usage

```sh
m365 spo folder get [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder is located

`-f, --url [url]`
: The server-relative URL of the folder to retrieve. Specify either `folderUrl` or `id` but not both

`-i, --id [id]`
: The UniqueId (GUID) of the folder to retrieve. Specify either `url` or `id` but not both

`--withPermissions`
: Set if you want to return associated roles and permissions of the folder. 

--8<-- "docs/cmd/_global.md"

## Remarks

If no folder exists at the specified URL, you will get a `Please check the folder URL. Folder might not exist on the specified URL` error.

If root level folder is passed, you will get a `Please ensure the specified folder URL or folder Id does not refer to a root folder. Use \'spo list get\' with withPermissions instead' error.` Please use the command 'spo list get'.

## Examples

Get folder properties for folder with server-relative url _'/Shared Documents'_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder get --webUrl https://contoso.sharepoint.com/sites/project-x --url "/Shared Documents"
```

Get folder properties for folder with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder get --webUrl https://contoso.sharepoint.com/sites/project-x --id "b2307a39-e878-458b-bc90-03bc578531d6"
```

Get folder properties for folder with server-relative url _'/Shared Documents/Test1'_ located in site _https://contoso.sharepoint.com/sites/test

```sh
m365 spo folder get --webUrl https://contoso.sharepoint.com/sites/test --url "Shared Documents/Test1" --withPermissions
```
