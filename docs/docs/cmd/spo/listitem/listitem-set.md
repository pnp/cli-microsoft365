# spo listitem set

Updates a list item in the specified list

## Usage

```sh
m365 spo listitem set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the item should be updated

`-i, --id <id>`
: ID of the list item to update.

`-l, --listId [listId]`
: ID of the list where the item should be updated. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: Title of the list where the item should be updated. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list where the item should be updated. Specify either `listTitle`, `listId` or `listUrl`

`-c, --contentType [contentType]`
: The name or the ID of the content type to associate with the updated item

`-s, --systemUpdate`
: Update the item without updating the modified date and modified by fields

--8<-- "docs/cmd/_global.md"

## Remarks

!!! tip "When using DateTime fields"
    When updating a list item with a DateTime field, use the timezone and the format that the site expects, based on its regional settings. Alternatively, a format which works on all regions is the following: `yyyy-MM-dd HH:mm:ss`. However, you should use the local timezone in all situations. UTC date/time or ISO 8601 formatted date/time is not supported.

## Examples

Update an item with id _147_ with title _Demo Item_ and content type name _Item_ in list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --contentType Item --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Item"
```

Update an item with id _147_ with title _Demo Multi Managed Metadata Field_ and a single-select metadata field named _SingleMetadataField_ in list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Managed Metadata Field" --SingleMetadataField "TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377;"
```

Update an item with id _147_ with Title _Demo Multi Managed Metadata Field_ and a multi-select metadata field named _MultiMetadataField_ in list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Managed Metadata Field" --MultiMetadataField "TermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75;"
```

Update an item with id 147 with Title _Demo Single Person Field_ and a single-select people field named _SinglePeopleField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Person Field" --SinglePeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'}]"
```

Update an item with id _147_ with Title _Demo Multi Person Field_ and a multi-select people field named _MultiPeopleField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Person Field" --MultiPeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'},{'Key':'i:0#.f|membership|john.doe@conotoso.com'}]"
```

Update the field _Title_ and _CustomHyperlink_ of an item with a specific id in a list retrieved by server-relative URL in a specific site

```sh
m365 spo listitem set --listUrl '/sites/project-x/lists/Demo List' --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Hyperlink Field" --CustomHyperlink "https://www.bing.com, Bing"
```

Update an item with a specific Title and multi-choice value

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo multi-choice Field" --MultiChoiceField "Choice 1;#Choice 2;#Choice 3"
```

Update an item with a specific Title and DateTime value

```sh
m365 spo listitem set --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo DateTime Field" --SomeDateTimeField "2023-01-16 15:30:00"
```

## Response

=== "JSON"

    ```json
    {
      "FileSystemObjectType": 0,
      "Id": 236,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ID": 236,
      "ContentTypeId": "0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803",
      "Title": "Updated Title",
      "Modified": "2022-11-16T21:10:06Z",
      "Created": "2022-11-16T20:56:31Z",
      "AuthorId": 10,
      "EditorId": 10,
      "OData__UIVersionString": "7.0",
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
    ID                           : 236
    Id                           : 236
    Modified                     : 2022-11-16T21:10:37Z
    OData__UIVersionString       : 8.0
    OData__vti_ItemDeclaredRecord: null
    ServerRedirectedEmbedUri     : null
    ServerRedirectedEmbedUrl     :
    Title                        : Updated Title
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ID,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,ComplianceAssetId,OData__vti_ItemDeclaredRecord
    0,236,,,236,0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803,Updated Title,2022-11-16T21:10:55Z,2022-11-16T20:56:31Z,10,10,9.0,1,cac57513-e870-4e7a-9f23-f4ea10e14f4e,,
    ```

=== "Markdown"

    ```md
    # spo listitem set --listTitle "My List" --id "236" --webUrl "https://contoso.sharepoint.com" --Title "Updated Title"

    Date: 2/20/2023

    ## Updated Title (236)

    Property | Value
    ---------|-------
    FileSystemObjectType | 0
    Id | 236
    ServerRedirectedEmbedUri | null
    ServerRedirectedEmbedUrl |
    ContentTypeId | 0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803
    Title | Updated Title
    ComplianceAssetId | null
    FieldName1 | null
    ID | 236
    Modified | 2022-11-16T21:10:06Z
    Created | 2022-11-16T20:56:31Z
    AuthorId | 10
    EditorId | 10
    OData\_\_UIVersionString | 7.0
    Attachments | true
    GUID | cac57513-e870-4e7a-9f23-f4ea10e14f4e
    OData__vti_ItemDeclaredRecord | null
    ```
