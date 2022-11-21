# spo listitem list

Gets a list of items from the specified list

## Usage

```sh
m365 spo listitem list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site from which the item should be retrieved.

`-i, --listId [listId]`
: ID of the list to retrieve items from. Specify `listId` or `listTitle` but not both.

`-l, --listId [listId]`
: ID of the list where the item should be added. Specify either `listTitle`, `listId`, or `listUrl` but not multiple.

`-t, --listTitle [listTitle]`
: Title of the list where the item should be added. Specify either `listTitle`, `listId`, or `listUrl` but not multiple.

`-q, --camlQuery [camlQuery]`
: CAML query to use to query the list of items with.

`-f, --fields [fields]`
: Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested. Specify `camlQuery` or `fields` but not both.

`-l, --filter [filter]`
: OData filter to use to query the list of items with. Specify `camlQuery` or `filter` but not both.

`-p, --pageSize [pageSize]`
: Number of list items to return. Specify `camlQuery` or `pageSize` but not both.

`-n, --pageNumber [pageNumber]`
: Page number to return if `pageSize` is specified (first page is indexed as value of 0).

--8<-- "docs/cmd/_global.md"

## Remarks

`pageNumber` is specified as a 0-based index. A value of `2` returns the third page of items.

If you want to specify a lookup type in the `properties` option, define which columns from the related list should be returned.

## Examples

Get all items from a list named Demo List

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x
```

From a list named _Demo List_ get all items with title _Demo list item_ using a CAML query

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --camlQuery "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo list item</Value></Eq></Where></Query></View>"
```

Get all items from a list with ID _935c13a0-cc53-4103-8b48-c1d0828eaa7f_

```sh
m365 spo listitem list --listId 935c13a0-cc53-4103-8b48-c1d0828eaa7f --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get all items from list named _Demo List_. For each item, retrieve the value of the _ID_, _Title_ and _Modified_ fields

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --fields "ID,Title,Modified"
```

Get all items from list named _Demo List_. For each item, retrieve the value of the _ID_, _Title_, _Modified_ fields, and the value of lookup field _Company_

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --fields "ID,Title,Modified,Company/Title"
```

From a list named _Demo List_ get all items with title _Demo list item_ using an OData filter

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --filter "Title eq 'Demo list item'"
```

From a list named _Demo List_ get the second batch of 10 items

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --pageSize 10 --pageNumber 2
```

Get all items from a list by server-relative URL

```sh
m365 spo listitem list --listUrl /sites/project-x/documents --webUrl https://contoso.sharepoint.com/sites/project-x
```

## Response

=== "JSON"

    ```json
    [
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
    ]
    ```

=== "Text"

    ```text
    Id   Title
    ---  -----
    236  Test
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,ComplianceAssetId,OData__vti_ItemDeclaredRecord
    0,236,,,0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803,Test,2022-11-16T21:00:03Z,2022-11-16T20:56:31Z,10,10,6.0,1,cac57513-e870-4e7a-9f23-f4ea10e14f4e,,
    ```
