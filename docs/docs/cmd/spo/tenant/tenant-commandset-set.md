# spo tenant commandset set

Update a ListView Command Set that is installed tenant wide.

## Usage

```sh
spo tenant commandset set [options]
```

## Options

`-i, --id <id>`
: The id of the ListView Command Set

`-t, --newTitle [newTitle]`
: The updated title of the ListView Command Set

`-l, --listType [listType]`
: The list or library type to register the ListView Command Set on. Allowed values `List` or `Library`.

`-i, --clientSideComponentId  [clientSideComponentId]`
: The Client Side Component Id (GUID) of the ListView Command Set.

`-p, --clientSideComponentProperties  [clientSideComponentProperties]`
: The Client Side Component properties of the ListView Command Set.

`-w, --webTemplate [webTemplate]`
: Optionally add a web template (e.g. STS#3, SITEPAGEPUBLISHING#0, etc) as a filter for what kind of sites the ListView Command Set is registered on.

`--location [location]`
: The location of the ListView Command Set. Allowed values `ContextMenu`, `CommandBar` or `Both`. Defaults to `CommandBar`.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! warning "Escaping JSON in PowerShell"
    When using the `--clientSideComponentProperties` option it's possible to enter a JSON string. In PowerShell 5 to 7.2 [specific escaping rules](./../../../user-guide/using-cli.md#escaping-double-quotes-in-powershell) apply due to an issue. Remember that you can also use [file tokens](./../../../user-guide/using-cli.md#passing-complex-content-into-cli-options) instead.

## Examples

Updates the title of a ListView Command Set that's deployed tenant wide.

```sh
m365 spo tenant commandset set --id 4  --newTitle "Some customizer"
```

Updates the properties of a ListView Command Set.

```sh
m365 spo tenant commandset  set --id 3  --clientSideComponentProperties '{ "someProperty": "Some value" }'
```

## Response

=== "JSON"

    ```json
    {
      "FileSystemObjectType": 0,
      "Id": 4,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ContentTypeId": "0x00693E2C487575B448BD420C12CEAE7EFE",
      "Title": "Some ListView Command Set",
      "Modified": "2023-01-11T15:47:38Z",
      "Created": "2023-01-11T15:47:38Z",
      "AuthorId": 9,
      "EditorId": 9,
      "OData__UIVersionString": "1.0",
      "Attachments": false,
      "GUID": "14125658-a9bc-4ddf-9c75-1b5767c9a337",
      "ComplianceAssetId": null,
      "TenantWideExtensionComponentId": "7096cded-b83d-4eab-96f0-df477ed7c0bc",
      "TenantWideExtensionComponentProperties": "{\"testMessage\":\"Test message\"}",
      "TenantWideExtensionWebTemplate": null,
      "TenantWideExtensionListTemplate": 101,
      "TenantWideExtensionLocation": "ClientSideExtension.ListViewCommandSet.ContextMenu",
      "TenantWideExtensionSequence": 0,
      "TenantWideExtensionHostProperties": null,
      "TenantWideExtensionDisabled": false
    }
    ```

=== "Text"

    ```text
    Attachments                           : false
    AuthorId                              : 9
    ComplianceAssetId                     : null
    ContentTypeId                         : 0x00693E2C487575B448BD420C12CEAE7EFE
    Created                               : 2023-01-11T15:47:38Z
    EditorId                              : 9
    FileSystemObjectType                  : 0
    GUID                                  : 14125658-a9bc-4ddf-9c75-1b5767c9a337
    Id                                    : 4
    Modified                              : 2023-01-11T15:47:38Z
    OData__UIVersionString                : 1.0
    ServerRedirectedEmbedUri              : null
    ServerRedirectedEmbedUrl              :
    TenantWideExtensionComponentId        : 7096cded-b83d-4eab-96f0-df477ed7c0bc
    TenantWideExtensionComponentProperties: {"testMessage":"Test message"}
    TenantWideExtensionDisabled           : false
    TenantWideExtensionHostProperties     : null
    TenantWideExtensionListTemplate       : 101
    TenantWideExtensionLocation           : ClientSideExtension.ListViewCommandSet.ContextMenu
    TenantWideExtensionSequence           : 0
    TenantWideExtensionWebTemplate        : null
    Title                                 : Some ListView Command Set
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,ComplianceAssetId,TenantWideExtensionComponentId,TenantWideExtensionComponentProperties,TenantWideExtensionWebTemplate,TenantWideExtensionListTemplate,TenantWideExtensionLocation,TenantWideExtensionSequence,TenantWideExtensionHostProperties,TenantWideExtensionDisabled
    0,4,,,0x00693E2C487575B448BD420C12CEAE7EFE,Some ListView Command Set,2023-01-11T15:47:38Z,2023-01-11T15:47:38Z,9,9,1.0,,14125658-a9bc-4ddf-9c75-1b5767c9a337,,7096cded-b83d-4eab-96f0-df477ed7c0bc,"{""testMessage"":""Test message""}",,101,ClientSideExtension.ListViewCommandSet.ContextMenu,0,,
    ```

=== "Markdown"

    ```md
    # spo tenant commandset set --id 4 --newTitle "Some ListView Command Set"

    Date: 20/05/2023

    ## Some ListView Command Set (4)

    Property | Value
    ---------|-------
    FileSystemObjectType | 0
    Id | 4
    ServerRedirectedEmbedUri | null
    ServerRedirectedEmbedUrl |
    ContentTypeId | 0x00693E2C487575B448BD420C12CEAE7EFE
    Title | Some customizer
    Modified | 2023-01-11T15:47:38Z
    Created | 2023-01-11T15:47:38Z
    AuthorId | 9
    EditorId | 9
    OData\_\_UIVersionString | 1.0
    Attachments | false
    GUID | 14125658-a9bc-4ddf-9c75-1b5767c9a337
    ComplianceAssetId | null
    TenantWideExtensionComponentId | 7096cded-b83d-4eab-96f0-df477ed7c0bc
    TenantWideExtensionComponentProperties | {"testMessage":"Test message"}
    TenantWideExtensionWebTemplate | null
    TenantWideExtensionListTemplate | 101
    TenantWideExtensionLocation | ClientSideExtension.ListViewCommandSet.ContextMenu
    TenantWideExtensionSequence | 0
    TenantWideExtensionHostProperties | null
    TenantWideExtensionDisabled | false
    ```
