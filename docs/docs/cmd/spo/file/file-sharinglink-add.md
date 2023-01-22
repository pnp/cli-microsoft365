# spo file sharinglink add

Creates a new sharing link for a file

## Usage

```sh
m365 spo file sharinglink add [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located.

`--fileUrl [fileUrl]`
: The server-relative (decoded) URL of the file. Specify either `fileUrl` or `fileId` but not both.

`--fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both.

`--type <type>`
: The type of sharing link to create. Either `view` or `edit`.

`--expirationDateTime [expirationDateTime]`
: The date and time to set the expiration. This should be defined as a valid ISO 8601 string.

`--scope [scope]`
: The scope of link to create. Either `anonymous` or `organization`. If not specified, the default of the organization will be used.

--8<-- "docs/cmd/_global.md"

## Examples

Creates a sharing link of a specific type for a file by id

```sh
m365 spo file sharinglink add --webUrl https://contoso.sharepoint.com --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --type view
```

Creates a sharing link of a specific type for a file by url

```sh
m365 spo file sharinglink add --webUrl https://contoso.sharepoint.com --fileUrl "Shared Documents/Test1.docx" --type edit
```

Creates a sharing link of a file by url with type, scope and expirationDateTime parameter

```sh
m365 spo file sharinglink add --webUrl https://contoso.sharepoint.com --fileUrl "Shared Documents/Test1.docx" --type edit --scope anonymous --expirationDateTime "2023-01-09T16:20:00Z"
```

## Response

=== "JSON"

    ```json
    {
      "id": "1e581e93-609e-4077-8152-c43865db684c",
      "roles": [
        "read"
      ],
      "expirationDateTime": "2023-10-01T07:00:00Z",
      "hasPassword": false,
      "link": {
        "scope": "anonymous",
        "type": "view",
        "webUrl": "https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBSyMcgRPtyAxqqNAeiEp1kg",
        "preventsDownload": false
      }
    }
    ```

=== "Text"

    ```text
    hasPassword       : false
    expirationDateTime: 2023-10-01T07:00:00Z
    id                : 1e581e93-609e-4077-8152-c43865db684c
    link              : {"scope":"anonymous","type":"view","webUrl":"https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBSyMcgRPtyAxqqNAeiEp1kg","preventsDownload":false}
    roles             : ["read"]
    ```

=== "CSV"

    ```csv
    id,expirationDateTime,roles,hasPassword,link
    1e581e93-609e-4077-8152-c43865db684c,2023-10-01T07:00:00Z,"[""read""]",,"{""scope"":""anonymous"",""type"":""view"",""webUrl"":""https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBSyMcgRPtyAxqqNAeiEp1kg"",""preventsDownload"":false}"
    ```
