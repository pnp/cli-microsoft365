# spo file version list

Retrieves all versions of a file

## Usage

```sh
m365 spo file version list [options]
```

## Options

`-w, --webUrl <webUrl>`
: The URL of the site where the file is located

`-u, --fileUrl [fileUrl]`
: The site-relative or server-relative URL of the file. Specify either `fileUrl` or `fileId` but not both

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both

--8<-- "docs/cmd/_global.md"

## Examples

List file versions of a specific file based on the ID of the file

```sh
m365 spo file version list --webUrl https://contoso.sharepoint.com --fileId 'b2307a39-e878-458b-bc90-03bc578531d6'
```

List file versions of a specific file based on the site-relative URL of the file

```sh
m365 spo file version list --webUrl https://contoso.sharepoint.com --fileUrl '/Shared Documents/Document.docx'
```

List file versions of a specific file based on the server-relative URL of the file

```sh
m365 spo file version list --webUrl https://contoso.sharepoint.com/sites/project-x --fileUrl '/sites/project-x/Shared Documents/Document.docx'
```

## Response

=== "JSON"

    ```json
    [
      {
        "CheckInComment": "",
        "Created": "2022-10-30T12:03:06Z",
        "ID": 512,
        "IsCurrentVersion": false,
        "Length": "18898",
        "Size": 18898,
        "Url": "_vti_history/512/Shared Documents/Document.docx",
        "VersionLabel": "1.0"
      }
    ]
    ```

=== "Text"

    ```text
    Created               ID    IsCurrentVersion  VersionLabel
    --------------------  ----  ----------------  ------------
    2022-10-30T12:03:06Z  512   false             1.0
    ```

=== "CSV"

    ```csv
    Created,ID,IsCurrentVersion,VersionLabel
    2022-10-30T12:03:06Z,512,false,1.0
    ```
