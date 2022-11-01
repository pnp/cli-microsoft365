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

`--clientSideComponentProperties [clientSideComponentProperties]`
: The Client Side Component properties of the application customizer.

`--webTemplate [webTemplate]`
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

=== "CSV"

    ```csv
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
