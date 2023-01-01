# spo file sharinglink get

Gets details about a specific sharing link

## Usage

```sh
m365 spo file sharinglink get [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located.

`--fileUrl [fileUrl]`
: The server-relative URL of the file. Specify either `fileUrl` or `fileId` but not both.

`--fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both.

`-i, --id <id>`
: The ID of the unique share.

--8<-- "docs/cmd/_global.md"

## Examples

Gets a specific sharing link of a file by id.

```sh
m365 spo file sharinglink get --webUrl 'https://contoso.sharepoint.com/sites/demo' --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --id U1BEZW1vIFZpc2l0b3Jz
```

Gets a specific sharing link of a file by a specified site-relative URL.

```sh
m365 spo file sharinglink get --webUrl 'https://contoso.sharepoint.com/sites/demo' --fileUrl 'Shared Documents/document.docx' --id U1BEZW1vIFZpc2l0b3Jz
```

Gets a specific sharing link of a file by a specified server-relative URL.

```sh
m365 spo file sharinglink get --webUrl 'https://contoso.sharepoint.com/sites/demo' --fileUrl '/sites/demo/Shared Documents/document.docx' --id U1BEZW1vIFZpc2l0b3Jz
```

## Response

=== "JSON"

    ```json
    {
      "id": "U1BEZW1vIFZpc2l0b3Jz",
      "roles": [
        "read"
      ],
      "grantedToV2": {
        "siteGroup": {
          "displayName": "Demo Visitors",
          "id": "5",
          "loginName": "Demo Visitors"
        }
      },
      "grantedTo": {
        "user": {
          "displayName": "Demo Visitors"
        }
      },
      "inheritedFrom": {}
    }
    ```

=== "Text"

    ```text
    grantedTo    : {"user":{"displayName":"Demo Visitors"}}
    grantedToV2  : {"siteGroup":{"displayName":"Demo Visitors","id":"5","loginName":"Demo Visitors"}}
    id           : U1BEZW1vIFZpc2l0b3Jz
    inheritedFrom: {}
    roles        : ["read"]
    ```

=== "CSV"

    ```csv
    id,roles,grantedToV2,grantedTo,inheritedFrom
    U1BEZW1vIFZpc2l0b3Jz,"[""read""]","{""siteGroup"":{""displayName"":""Demo Visitors"",""id"":""5"",""loginName"":""Demo Visitors""}}","{""user"":{""displayName"":""Demo Visitors""}}",{}
    ```

=== "Markdown"

    ```md
    # spo file sharinglink get --webUrl "https://contoso.sharepoint.com/sites/demo" --fileUrl "/sites/demo/Shared Documents/document.docx" --id "U1BEZW1vIFZpc2l0b3Jz"

    Date: 1/1/2023

    ## undefined (U1BEZW1vIFZpc2l0b3Jz)

    Property | Value
    ---------|-------
    id | U1BEZW1vIFZpc2l0b3Jz
    roles | ["read"]
    grantedToV2 | {"siteGroup":{"displayName":"Demo Visitors","id":"5","loginName":"Demo Visitors"}}
    grantedTo | {"user":{"displayName":"Demo Visitors"}}
    inheritedFrom | {}
    ```
