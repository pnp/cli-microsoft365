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

`-r, --recursive`
: Set to retrieve nested folders

--8<-- "docs/cmd/_global.md"

## Examples

Gets list of folders under a parent folder with site-relative url _/Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents'
```

Gets recursive list of folders under a specific folder on a specific site

```sh
m365 spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents' --recursive
```

## Response

=== "JSON"

    ```json
    [  
      {
        "Exists": true,
        "IsWOPIEnabled": false,
        "ItemCount": 9,
        "Name": "Folder A",
        "ProgID": null,
        "ServerRelativeUrl": "/Shared Documents/Folder A",
        "TimeCreated": "2022-04-26T12:30:56Z",
        "TimeLastModified": "2022-04-26T12:50:14Z",
        "UniqueId": "20523746-971b-4488-aa6d-b45d645f61c5",
        "WelcomePage": ""
      }
    ]
    ```

=== "Text"

    ```text
    Name     ServerRelativeUrl
    -------  -------------------------
    Folder A /Shared Documents/Folder A
    ```

=== "CSV"

    ```csv
    Name,ServerRelativeUrl
    Folder A,/Shared Documents/Folder A
    ```
