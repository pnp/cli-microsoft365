# spo applicationcustomizer add

Add an application customizer to a site. 

## Usage

```sh
m365 spo applicationcustomizer add [options]
```

## Options

`-t, --title <title>`
: The title of the Application Customizer

`-u, --webUrl [webUrl]`
: The site to add the Application Customizer on. If not specified, the application customizer is deployed tenant wide.

`-i, --clientSideComponentId <clientSideComponentId>`
: The Client Side Component Id (GUID) of the application customizer.

`-p, --clientSideComponentProperties [clientSideComponentProperties]`
: The Client Side Component properties of the application customizer.

`-w, --webTemplate [webTemplate]`
: Optionally add a web template (e.g. STS#3, SITEPAGEPUBLISHING#0, etc) when deploying tenant wide as a filter for what kind of sites the application customizer is registered on. 

--8<-- "docs/cmd/_global.md"

## Examples

Add an application customizer to the sales site.

```sh
m365 spo applicationcustomizer add --title "Some customizer" --clientSideComponentId 6b2a54c5-3317-49eb-8621-1bbb76263629 --webUrl https://contoso.sharepoint.com/sites/sales
```

Deploy an application customizer to all communication sites

```sh
m365 spo applicationcustomizer add --title "Some customizer" --clientSideComponentId 6b2a54c5-3317-49eb-8621-1bbb76263629 --webTemplate "SITEPAGEPUBLISHING#0"
```

## Response

The command will only return a response when adding a tenant-wide extension. 

=== "JSON"

    ```json
    {
      "FileSystemObjectType": 0,
      "Id": 21,
      "ServerRedirectedEmbedUri": null,
      "ServerRedirectedEmbedUrl": "",
      "ID": 21,
      "ContentTypeId": "0x0089AA3D6048D9A1418E2BE8FE410E851B",
      "Title": "Some customizer",
      "Modified": "2022-11-05T15:22:08Z",
      "Created": "2022-11-05T15:22:08Z",
      "AuthorId": 9,
      "EditorId": 9,
      "OData__UIVersionString": "1.0",
      "Attachments": false,
      "GUID": "3844156c-0ae5-464d-80a5-3195f8e78a8c",
      "ComplianceAssetId": null,
      "TenantWideExtensionComponentId": "6b2a54c5-3317-49eb-8621-1bbb76263629",
      "TenantWideExtensionComponentProperties": null,
      "TenantWideExtensionWebTemplate": "STS#0",
      "TenantWideExtensionListTemplate": 0,
      "TenantWideExtensionLocation": "ClientSideExtension.ApplicationCustomizer",
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
    ContentTypeId                         : 0x0089AA3D6048D9A1418E2BE8FE410E851B
    Created                               : 2022-11-25T10:53:35Z
    EditorId                              : 9
    FileSystemObjectType                  : 0
    GUID                                  : 8e1fd962-a926-4f4d-a67b-50fb8fa14f74
    ID                                    : 40
    Id                                    : 40
    Modified                              : 2022-11-25T10:53:35Z
    OData__UIVersionString                : 1.0
    ServerRedirectedEmbedUri              : null
    ServerRedirectedEmbedUrl              :
    TenantWideExtensionComponentId        : 6b2a54c5-3317-49eb-8621-1bbb76263629
    TenantWideExtensionComponentProperties: null
    TenantWideExtensionDisabled           : false
    TenantWideExtensionHostProperties     : null
    TenantWideExtensionListTemplate       : 0
    TenantWideExtensionLocation           : ClientSideExtension.ApplicationCustomizer
    TenantWideExtensionSequence           : 0
    TenantWideExtensionWebTemplate        : STS#0
    Title                                 : Some customizer
    ```

=== "CSV"

    ```csv
    FileSystemObjectType,Id,ServerRedirectedEmbedUri,ServerRedirectedEmbedUrl,ID,ContentTypeId,Title,Modified,Created,AuthorId,EditorId,OData__UIVersionString,Attachments,GUID,ComplianceAssetId,TenantWideExtensionComponentId,TenantWideExtensionComponentProperties,TenantWideExtensionWebTemplate,TenantWideExtensionListTemplate,TenantWideExtensionLocation,TenantWideExtensionSequence,TenantWideExtensionHostProperties,TenantWideExtensionDisabled
    0,41,,,41,0x0089AA3D6048D9A1418E2BE8FE410E851B,Some customizer,2022-11-25T10:53:58Z,2022-11-25T10:53:58Z,9,9,1.0,,5f10e339-ab9e-4e36-9881-1f42a6b82c09,,6b2a54c5-3317-49eb-8621-1bbb76263629,,STS#0,0,ClientSideExtension.ApplicationCustomizer,0,,
    ```
