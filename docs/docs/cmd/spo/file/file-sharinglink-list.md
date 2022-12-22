# spo file sharinglink list

Lists all the sharing links of a specific file

## Usage

```sh
m365 spo file sharinglink list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located

`-f, --fileUrl [fileUrl]`
: The server-relative (decoded) URL of the file. Specify either `fileUrl` or `fileId` but not both

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both

--8<-- "docs/cmd/_global.md"

## Examples

List sharing links of a file by id

```sh
m365 spo file sharinglink list --webUrl https://contoso.sharepoint.com/sites/demo --fileId daebb04b-a773-4baa-b1d1-3625418e3234
```

List sharing links of a file by url

```sh
m365 spo file sharinglink list --webUrl https://contoso.sharepoint.com/sites/demo --fileUrl /sites/demo/shared%20documents/document.docx
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "2a021f54-90a2-4016-b3b3-5f34d2e7d932",
        "roles": [
          "read"
        ],
        "hasPassword": false,
        "grantedToIdentitiesV2": [
          {
            "user": {
              "displayName": "John Doe",
              "email": "john@contoso.onmicrosoft.com",
              "id": "fe36f75e-c103-410b-a18a-2bf6df06ac3a"
            },
            "siteUser": {
              "displayName": "John Doe",
              "email": "john@contoso.onmicrosoft.com",
              "id": "9",
              "loginName": "i:0#.f|membership|john@contoso.onmicrosoft.com"
            }
          }
        ],
        "grantedToIdentities": [ 
          {
            "user": {
              "displayName": "John Doe",
              "email": "john@contoso.onmicrosoft.com",
              "id": "fe36f75e-c103-410b-a18a-2bf6df06ac3a"
            }
          }
        ],
        "link": {
          "scope": "anonymous",
          "type": "view",
          "webUrl": "https://contoso.sharepoint.com/:b:/s/demo/EY50lub3559MtRKfj2hrZqoBWnHOpGIcgi4gzw9XiWYJ-A",
          "preventsDownload": false
        }
      }
    ]
    ```

=== "Text"

    ```text
    id                                    roles  link
    ------------------------------------  -----  -----------------------------------------------------------------------------------------------------------
    2a021f54-90a2-4016-b3b3-5f34d2e7d932  read   https://contoso.sharepoint.com/:b:/s/demo/EY50lub3559MtRKfj2hrZqoBWnHOpGIcgi4gzw9XiWYJ-A
    ```

=== "CSV"

    ```csv
    id,roles,link
    2a021f54-90a2-4016-b3b3-5f34d2e7d932,read,https://contoso.sharepoint.com/:b:/s/demo/EY50lub3559MtRKfj2hrZqoBWnHOpGIcgi4gzw9XiWYJ-A
    ```
