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
: The server-relative URL of the file. Specify either `fileUrl` or `fileId` but not both

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both

--8<-- "docs/cmd/_global.md"

## Examples

Get file versions in a specific site based on fileUrl

```sh
m365 spo file version list --webUrl https://contoso.sharepoint.com --fileId 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Get file versions in a specific site based on fileId

```sh
m365 spo file version list --webUrl https://contoso.sharepoint.com --fileUrl '/Shared Documents/Document.docx'
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
      },
      {
        "CheckInComment": "",
        "Created": "2022-10-30T12:06:13Z",
        "ID": 1024,
        "IsCurrentVersion": false,
        "Length": "21098",
        "Size": 21098,
        "Url": "_vti_history/1024/Shared Documents/Document.docx",
        "VersionLabel": "2.0"
      }
    ]
    ```

=== "Text"

    ```text
    CheckInComment  Created               ID    IsCurrentVersion  Length  Size   Url                                               VersionLabel
    --------------  --------------------  ----  ----------------  ------  -----  ------------------------------------------------  ------------
                    2022-10-30T12:03:06Z  512   false             18898   18898  _vti_history/512/Shared Documents/Document.docx   1.0
                    2022-10-30T12:06:13Z  1024  false             21098   21098  _vti_history/1024/Shared Documents/Document.docx  2.0
    ```

=== "CSV"

    ```csv
    CheckInComment,Created,ID,IsCurrentVersion,Length,Size,Url,VersionLabel
    ,2022-10-30T12:03:06Z,512,,18898,18898,_vti_history/512/Shared Documents/Document.docx,1.0
    ,2022-10-30T12:06:13Z,1024,,21098,21098,_vti_history/1024/Shared Documents/Document.docx,2.0
    ```

