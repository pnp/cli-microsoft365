# spo file sharinglink add

Creates a new sharing link to a file

## Usage

```sh
m365 spo file sharinglink add [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located.

`-f, --fileUrl [fileUrl]`
: The server-relative (decoded) URL of the file. Specify either `fileUrl` or `fileId` but not both.

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Creates a sharing link of a file by id and type parameter

```sh
m365 spo file sharinglink add --webUrl https://contoso.sharepoint.com/sites/demo --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --type view
```

Creates a sharing link of a file by url and type parameter

```sh
m365 spo file sharinglink add --webUrl https://contoso.sharepoint.com/sites/demo --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --type edit
```

Creates a sharing link of a file by url with type, scope and expirationDateTime parameter

```sh
m365 spo file sharinglink add --webUrl https://contoso.sharepoint.com/sites/demo --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --type edit --scope anonymous --expirationDateTime "2023-01-09T16:20:00Z"
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
    id   : 1e581e93-609e-4077-8152-c43865db684c
    link : https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBSyMcgRPtyAxqqNAeiEp1kg
    roles: read
    scope: anonymous
    ```

=== "CSV"

    ```csv
    id,roles,link,scope
    1e581e93-609e-4077-8152-c43865db684c,read,https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBSyMcgRPtyAxqqNAeiEp1kg,anonymous
    ```

=== "Markdown"

    ```md
    # spo file sharinglink add --webUrl "https://contoso.sharepoint.com" --fileUrl "/Shared Documents/Document.docx" --type "view" --scope "anonymous" --expirationDateTime "2023-10-01"

    Date: 9/1/2023

    ## undefined (3374c33e-8f13-4c2f-9b42-e7450786647f)

    Property | Value
    ---------|-------
    id | 1e581e93-609e-4077-8152-c43865db684c
    roles | ["read"]
    expirationDateTime | 2023-10-01T07:00:00Z
    hasPassword | false
    link | {"scope":"anonymous","type":"view","webUrl":"https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBSyMcgRPtyAxqqNAeiEp1kg","preventsDownload":false}
    ```
