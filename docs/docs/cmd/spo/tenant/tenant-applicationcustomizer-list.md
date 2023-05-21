# spo tenant applicationcustomizer list

Get a list of application customizers that are installed tenant wide

## Usage

```sh
spo tenant applicationcustomizer list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Retrieves a list of application customizers.

```sh
m365 spo tenant applicationcustomizer list
```

## Response

=== "JSON"

    ```json
    [
      {
        "FileSystemObjectType": 0,
        "Id": 8,
        "ServerRedirectedEmbedUri": null,
        "ServerRedirectedEmbedUrl": "",
        "ID": 8,
        "ContentTypeId": "0x00693E2C487575B448BD420C12CEAE7EFE",
        "Title": "HelloWorld",
        "Modified": "2023-05-21T14:31:30Z",
        "Created": "2023-05-21T14:31:30Z",
        "AuthorId": 9,
        "EditorId": 9,
        "OData__UIVersionString": "1.0",
        "Attachments": false,
        "GUID": "23951a41-f613-440e-8119-8f1e87df1d1a",
        "OData__ColorTag": null,
        "ComplianceAssetId": null,
        "TenantWideExtensionComponentId": "d54e75e7-af4d-455f-9101-a5d906692ecd",
        "TenantWideExtensionComponentProperties": "{\"testMessage\":\"Test message\"}",
        "TenantWideExtensionWebTemplate": null,
        "TenantWideExtensionListTemplate": 0,
        "TenantWideExtensionLocation": "ClientSideExtension.ApplicationCustomizer",
        "TenantWideExtensionSequence": 0,
        "TenantWideExtensionHostProperties": null,
        "TenantWideExtensionDisabled": false
      }
    ]
    ```

=== "Text"

    ```text
    TenantWideExtensionComponentId: d54e75e7-af4d-455f-9101-a5d906692ecd
    TenantWideExtensionWebTemplate: null
    Title                         : HelloWorld
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUrl,ID,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,TenantWideExtensionComponentId,TenantWideExtensionComponentProperties,TenantWideExtensionListTemplate,TenantWideExtensionLocation,TenantWideExtensionSequence,TenantWideExtensionDisabled
    0,8,,8,0x00693E2C487575B448BD420C12CEAE7EFE,HelloWorld,2023-05-21T14:31:30Z,2023-05-21T14:31:30Z,9,9,1.0,,23951a41-f613-440e-8119-8f1e87df1d1a,d54e75e7-af4d-455f-9101-a5d906692ecd,"{""testMessage"":""Test message""}",0,ClientSideExtension.ApplicationCustomizer,0,
    ```

=== "Markdown"

    ```md
    # spo tenant applicationcustomizer list

    Date: 5/21/2023

    ## HelloWorld (8)

    Property | Value
    ---------|-------
    FileSystemObjectType | 0
    Id | 8
    ServerRedirectedEmbedUrl |
    ID | 8
    ContentTypeId | 0x00693E2C487575B448BD420C12CEAE7EFE
    Title | HelloWorld
    Modified | 2023-05-21T14:31:30Z
    Created | 2023-05-21T14:31:30Z
    AuthorId | 9
    EditorId | 9
    OData\_\_UIVersionString | 1.0
    Attachments | false
    GUID | 23951a41-f613-440e-8119-8f1e87df1d1a
    TenantWideExtensionComponentId | d54e75e7-af4d-455f-9101-a5d906692ecd
    TenantWideExtensionComponentProperties | {"testMessage":"Test message"}
    TenantWideExtensionListTemplate | 0
    TenantWideExtensionLocation | ClientSideExtension.ApplicationCustomizer
    TenantWideExtensionSequence | 0
    TenantWideExtensionDisabled | false
    ```
