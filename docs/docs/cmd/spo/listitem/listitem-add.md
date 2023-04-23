# spo listitem add

Creates a list item in the specified list

## Usage

```sh
m365 spo listitem add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the item should be added

`-l, --listId [listId]`
: ID of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: Title of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

`-c, --contentType [contentType]`
: The name or the ID of the content type to associate with the new item

`-f, --folder [folder]`
: The list-relative URL of the folder where the item should be created

--8<-- "docs/cmd/_global.md"

## Remarks

!!! warning "When using DateTime fields"
    When creating a list item with a DateTime field, use the timezone and the format that the site expects, based on its regional settings. Alternatively, a format which works on all regions is the following: `yyyy-MM-dd HH:mm:ss`. However, you should use the local timezone in all situations. UTC date/time or ISO 8601 formatted date/time is not supported.

## Examples

Add an item with Title _Demo Item_ and content type name _Item_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem add --contentType Item --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Item"
```

Add an item with Title _Demo Multi Managed Metadata Field_ and a single-select metadata field named _SingleMetadataField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem add --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Managed Metadata Field" --SingleMetadataField "TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377;"
```

Add an item with Title _Demo Multi Managed Metadata Field_ and a multi-select metadata field named _MultiMetadataField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem add --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Managed Metadata Field" --MultiMetadataField "TermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75;"
```

Add an item with Title _Demo Single Person Field_ and a single-select people field named _SinglePeopleField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem add --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Person Field" --SinglePeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'}]"
```

Add an item with Title _Demo Multi Person Field_ and a multi-select people field named _MultiPeopleField_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem add --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Person Field" --MultiPeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'},{'Key':'i:0#.f|membership|john.doe@conotoso.com'}]"
```

Add an item with Title _Demo Hyperlink Field_ and a hyperlink field named _CustomHyperlink_ to list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem add --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Hyperlink Field" --CustomHyperlink "https://www.bing.com, Bing"
```

Add an item with a specific title to a list retrieved by server-relative URL in a specific site

```sh
m365 spo listitem add --contentType Item --listUrl /sites/project-x/Documents --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Item"
```

Add an item with a specific Title and multi-choice value

```sh
m365 spo listitem add --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo multi-choice Field" --MultiChoiceField "Choice 1;#Choice 2;#Choice 3"
```

Add an item with a specific Title and DateTime value

```sh
m365 spo listitem add --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo DateTime Field" --SomeDateTimeField "2023-01-16 15:30:00"
```

## Response

=== "JSON"

    ```json
    {
      "FileSystemObjectType": 0,
      "Id": 234,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ID": 234,
      "ContentTypeId": "0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803",
      "Title": "Test",
      "Modified": "2022-11-16T20:55:45Z",
      "Created": "2022-11-16T20:55:45Z",
      "AuthorId": 10,
      "EditorId": 10,
      "OData__UIVersionString": "1.0",
      "Attachments": false,
      "GUID": "352e3855-56fa-4b68-b6be-4644d6adf204",
      "ComplianceAssetId": null,
      "OData__vti_ItemDeclaredRecord": null,
    }
    ```

=== "Text"

    ```text
    Attachments                  : false
    AuthorId                     : 10
    ComplianceAssetId            : null
    ContentTypeId                : 0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803
    Created                      : 2022-11-16T20:56:31Z
    EditorId                     : 10
    FileSystemObjectType         : 0
    GUID                         : cac57513-e870-4e7a-9f23-f4ea10e14f4e
    ID                           : 236
    Id                           : 236
    Modified                     : 2022-11-16T20:56:31Z
    OData__UIVersionString       : 1.0
    OData__vti_ItemDeclaredRecord: null
    ServerRedirectedEmbedUri     : null
    ServerRedirectedEmbedUrl     :
    Title                        : Test
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ID,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,ComplianceAssetId,OData__vti_ItemDeclaredRecord
    0,235,,,235,0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803,Test,2022-11-16T20:56:09Z,2022-11-16T20:56:09Z,10,10,1.0,,7aa8f3bd-a0a2-4974-81c8-2ac7ddc8e2d8,,
    ```

=== "Markdown"

    ```md
    # spo listitem add --contentType "Item" --listTitle "My List" --webUrl "https://contoso.sharepoint.com/sites/project-x" --Title "Test"

    Date: 2/20/2023

    ## Test (234)

    Property | Value
    ---------|-------
    FileSystemObjectType | 0
    Id | 234
    ServerRedirectedEmbedUri | null
    ServerRedirectedEmbedUrl |
    ContentTypeId | 0x01003CDBEB7138618C47A98D56499135D6EE0004C0F5794DEBCC4BAC981AC4AE1BD803
    Title | Test
    ComplianceAssetId | null
    FieldName1 | null
    ID | 234
    Modified | 2022-11-16T20:55:45Z
    Created | 2022-11-16T20:55:45Z
    AuthorId | 10
    EditorId | 10
    OData\_\_UIVersionString | 1.0
    Attachments | false
    GUID | 352e3855-56fa-4b68-b6be-4644d6adf204
    ```
