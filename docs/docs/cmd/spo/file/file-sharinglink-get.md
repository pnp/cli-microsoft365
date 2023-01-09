# spo file sharinglink get

Gets details about a specific sharing link of a file

## Usage

```sh
m365 spo file sharinglink get [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located.

`--fileUrl [fileUrl]`
: The server-relative (decoded) URL of the file. Specify either `fileUrl` or `fileId` but not both.

`--fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both.

`-i, --id <id>`
: The ID of the sharing link.

--8<-- "docs/cmd/_global.md"

## Examples

Gets a specific sharing link of a file by id.

```sh
m365 spo file sharinglink get --webUrl 'https://contoso.sharepoint.com/sites/demo' --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --id 1ba739c5-e693-4c16-9dfa-042e4ec62972
```

Gets a specific sharing link of a file by a specified site-relative URL.

```sh
m365 spo file sharinglink get --webUrl 'https://contoso.sharepoint.com/sites/demo' --fileUrl 'Shared Documents/document.docx' --id 1ba739c5-e693-4c16-9dfa-042e4ec62972
```

Gets a specific sharing link of a file by a specified server-relative URL.

```sh
m365 spo file sharinglink get --webUrl 'https://contoso.sharepoint.com/sites/demo' --fileUrl '/sites/demo/Shared Documents/document.docx' --id 1ba739c5-e693-4c16-9dfa-042e4ec62972
```

## Response

=== "JSON"

    ```json
    {
      "id": "1ba739c5-e693-4c16-9dfa-042e4ec62972",
      "roles": [
        "write"
      ],
      "hasPassword": false,
      "grantedToIdentitiesV2": [
        {
          "user": {
            "displayName": "John Doe",
            "email": "john@contoso.onmicrosoft.com",
            "id": "04355ecd-2124-4097-bc2b-c2295a71d7a3"
          },
          "siteUser": {
            "displayName": "John Doe",
            "email": "john@contoso.onmicrosoft.com",
            "id": "11",
            "loginName": "i:0#.f|membership|john@contoso.onmicrosoft.com"
          }
        }
      ],
      "grantedToIdentities": [
        {
          "user": {
            "displayName": "John Doe",
            "email": "john@contoso.onmicrosoft.com",
            "id": "04355ecd-2124-4097-bc2b-c2295a71d7a3"
          }
        }
      ],
      "link": {
        "scope": "organization",
        "type": "edit",
        "webUrl": "https://contoso.sharepoint.com/:w:/s/demo/EecoJa3lri9Hu9NWp-W0aBQB8ZqmGqA5tdIiaab4o-6BZw",
        "preventsDownload": false
      }
    }
    ```

=== "Text"

    ```text
    grantedToIdentities  : [{"user":{"displayName":"John Doe","email":"john@contoso.onmicrosoft.com","id":"04355ecd-2124-4097-bc2b-c2295a71d7a3"}}]
    grantedToIdentitiesV2: [{"user":{"displayName":"John Doe","email":"john@contoso.onmicrosoft.com","id":"04355ecd-2124-4097-bc2b-c2295a71d7a3"},"siteUser":{"displayName":"John Doe","email":"john@contoso.onmicrosoft.com","id":"11","loginName":"i:0#.f|membership|john@contoso.onmicrosoft.com"}}]
    hasPassword          : false
    id                   : 1ba739c5-e693-4c16-9dfa-042e4ec62972
    link                 : {"scope":"organization","type":"edit","webUrl":"https://contoso.sharepoint.com/:w:/s/demo/EecoJa3lri9Hu9NWp-W0aBQB8ZqmGqA5tdIiaab4o-6BZw","preventsDownload":false}
    roles                : ["write"]
    ```

=== "CSV"

    ```csv
    id,roles,hasPassword,grantedToIdentitiesV2,grantedToIdentities,link
    1ba739c5-e693-4c16-9dfa-042e4ec62972,"[""write""]",,"[{""user"":{""displayName"":""John Doe"",""email"":""john@contoso.onmicrosoft.com"",""id"":""04355ecd-2124-4097-bc2b-c2295a71d7a3""},""siteUser"":{""displayName"":""John Doe"",""email"":""john@contoso.onmicrosoft.com"",""id"":""11"",""loginName"":""i:0#.f|membership|john@contoso.onmicrosoft.com""}}]","[{""user"":{""displayName"":""John Doe"",""email"":""john@contoso.onmicrosoft.com"",""id"":""04355ecd-2124-4097-bc2b-c2295a71d7a3""}}]","{""scope"":""organization"",""type"":""edit"",""webUrl"":""https://contoso.sharepoint.com/:w:/s/demo/EecoJa3lri9Hu9NWp-W0aBQB8ZqmGqA5tdIiaab4o-6BZw"",""preventsDownload"":false}"
    ```
