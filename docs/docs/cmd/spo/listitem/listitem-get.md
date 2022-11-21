# spo listitem get

Gets a list item from the specified list

## Usage

```sh
m365 spo listitem get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site from which the item should be retrieved

`-i, --id <id>`
: ID of the item to retrieve.

`-l, --listId [listId]`
: ID of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: Title of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

`-p, --properties [properties]`
: Comma-separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested

--8<-- "docs/cmd/_global.md"

## Remarks

If you want to specify a lookup type in the `properties` option, define which columns from the related list should be returned.

## Examples

Get an item with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get an items _Title_ and _Created_ column with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Created"
```

Get an items _Title_, _Created_ column and lookup column _Company_ with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Created,Company/Title"
```

Get an item with specific properties from a list retrieved by server-relative URL in a specific site

```sh
m365 spo listitem get --listUrl /sites/project-x/documents --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Created,Company/Title"
```

## Response

=== "JSON"

    ```json
    {
      "FileSystemObjectType": 0,
      "Id": 236,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ContentTypeId": "0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803",
      "Title": "Test",
      "Modified": "2022-11-16T21:00:03Z",
      "Created": "2022-11-16T20:56:31Z",
      "AuthorId": 10,
      "EditorId": 10,
      "OData__UIVersionString": "6.0",
      "Attachments": true,
      "GUID": "cac57513-e870-4e7a-9f23-f4ea10e14f4e",
      "ComplianceAssetId": null,
      "OData__vti_ItemDeclaredRecord": null
    }
    ```

=== "Text"

    ```text
    Attachments                  : true
    AuthorId                     : 10
    ComplianceAssetId            : null
    ContentTypeId                : 0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803
    Created                      : 2022-11-16T20:56:31Z
    EditorId                     : 10
    FileSystemObjectType         : 0
    GUID                         : cac57513-e870-4e7a-9f23-f4ea10e14f4e
    Id                           : 236
    Modified                     : 2022-11-16T21:00:03Z
    OData__UIVersionString       : 6.0
    OData__vti_ItemDeclaredRecord: null
    ServerRedirectedEmbedUri     : null
    ServerRedirectedEmbedUrl     :
    Title                        : Test
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,ComplianceAssetId,OData__vti_ItemDeclaredRecord
    0,236,,,0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803,Test,2022-11-16T21:00:03Z,2022-11-16T20:56:31Z,10,10,6.0,1,cac57513-e870-4e7a-9f23-f4ea10e14f4e,,
    ```
