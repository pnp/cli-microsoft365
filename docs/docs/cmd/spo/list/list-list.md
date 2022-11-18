# spo list list

Gets all lists within the specified site

## Usage

```sh
m365 spo list list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the lists to retrieve are located

--8<-- "docs/cmd/_global.md"

## Examples

Return all lists located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list list --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)

## Response

=== "JSON"

    ```json
    [
      {
        "RootFolder": {
          "ServerRelativeUrl": "/_catalogs/theme"
        },
        "AllowContentTypes": true,
        "BaseTemplate": 123,
        "BaseType": 1,
        "ContentTypesEnabled": false,
        "CrawlNonDefaultViews": false,
        "Created": "2020-01-12T01:03:13Z",
        "CurrentChangeToken": {
          "StringValue": "1;3;66e5148c-7060-4479-88e7-636d79579148;638042267256930000;564174226"
        },
        "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000",
        "DefaultItemOpenUseListSetting": false,
        "Description": "Use the theme gallery to store themes. The themes in this gallery can be used by this site or any of its subsites.",
        "Direction": "none",
        "DisableCommenting": false,
        "DisableGridEditing": false,
        "DocumentTemplateUrl": null,
        "DraftVersionVisibility": 0,
        "EnableAttachments": false,
        "EnableFolderCreation": false,
        "EnableMinorVersions": false,
        "EnableModeration": false,
        "EnableRequestSignOff": true,
        "EnableVersioning": false,
        "EntityTypeName": "OData__x005f_catalogs_x002f_theme",
        "ExemptFromBlockDownloadOfNonViewableFiles": false,
        "FileSavePostProcessingEnabled": false,
        "ForceCheckout": false,
        "HasExternalDataSource": false,
        "Hidden": true,
        "Id": "66e5148c-7060-4479-88e7-636d79579148",
        "ImagePath": {
          "DecodedUrl": "/_layouts/15/images/itdl.png?rev=47"
        },
        "ImageUrl": "/_layouts/15/images/itdl.png?rev=47",
        "DefaultSensitivityLabelForLibrary": "",
        "IrmEnabled": false,
        "IrmExpire": false,
        "IrmReject": false,
        "IsApplicationList": false,
        "IsCatalog": true,
        "IsPrivate": false,
        "ItemCount": 41,
        "LastItemDeletedDate": "2020-01-12T01:03:13Z",
        "LastItemModifiedDate": "2020-01-12T01:03:18Z",
        "LastItemUserModifiedDate": "2020-01-12T01:03:18Z",
        "ListExperienceOptions": 0,
        "ListItemEntityTypeFullName": "SP.Data.OData__x005f_catalogs_x002f_themeItem",
        "MajorVersionLimit": 0,
        "MajorWithMinorVersionsLimit": 0,
        "MultipleDataList": false,
        "NoCrawl": true,
        "ParentWebPath": {
          "DecodedUrl": "/"
        },
        "ParentWebUrl": "/",
        "ParserDisabled": false,
        "ServerTemplateCanCreateFolders": false,
        "TemplateFeatureId": "00000000-0000-0000-0000-000000000000",
        "Title": "Theme Gallery",
        "Url": "/_catalogs/theme"
      }
    ]
    ```

=== "Text"

    ```text
    Id                                    Url
    ------------------------------------  ----------------
    66e5148c-7060-4479-88e7-636d79579148  /_catalogs/theme
    ```

=== "CSV"

    ```csv
    Id,Url
    Theme Gallery,/_catalogs/theme,66e5148c-7060-4479-88e7-636d79579148
    ```
