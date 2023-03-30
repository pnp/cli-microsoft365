# spo folder list

Returns all folders under the specified parent folder

## Usage

```sh
m365 spo folder list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folders to list are located.

`-p, --parentFolderUrl <parentFolderUrl>`
: The server- or site-relative URL of the parent folder.

`-f, --fields [fields]`
: Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested.

`--filter [filter]`
: OData filter to use to query the list of folders with.

`-r, --recursive`
: Set to retrieve nested folders.

--8<-- "docs/cmd/_global.md"

## Examples

Gets list of folders under a parent folder with site-relative URL

```sh
m365 spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents'
```

Gets recursive list of folders under a specific folder on a specific site

```sh
m365 spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/sites/project-x/Shared Documents' --recursive
```

Return a filtered list of folders and only return the list item ID

```sh
m365 spo folder list --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents' --fields ListItemAllFields/Id --filter "startswith(Name,'Folder')"
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
        "ServerRelativeUrl": "/sites/project-x/Shared Documents/Folder A",
        "TimeCreated": "2022-04-26T12:30:56Z",
        "TimeLastModified": "2022-04-26T12:50:14Z",
        "UniqueId": "20523746-971b-4488-aa6d-b45d645f61c5",
        "WelcomePage": ""
      }
    ]
    ```

=== "Text"

    ```text
    Name      ServerRelativeUrl
    --------  ------------------------------------------
    Folder A  /sites/project-x/Shared Documents/Folder A
    ```

=== "CSV"

    ```csv
    Name,ServerRelativeUrl
    Folder A,/sites/project-x/Shared Documents/Folder A
    ```

=== "Markdown"

    ```md
    # spo folder list --webUrl "https://contoso.sharepoint.com" --parentFolderUrl "/Shared Documents"

    Date: 29/3/2023

    ## Folder A (20523746-971b-4488-aa6d-b45d645f61c5)

    Property | Value
    ---------|-------
    Exists | true
    IsWOPIEnabled | false
    ItemCount | 9
    Name | Folder A
    ProgID | null
    ServerRelativeUrl | /sites/project-x/Shared Documents/Folder A
    TimeCreated | 2022-04-26T12:30:56Z
    TimeLastModified | 2022-04-26T12:50:14Z
    UniqueId | 20523746-971b-4488-aa6d-b45d645f61c5
    WelcomePage |
    ```
