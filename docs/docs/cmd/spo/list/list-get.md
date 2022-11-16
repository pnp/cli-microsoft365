# spo list get

Gets information about the specific list

## Usage

```sh
m365 spo list get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to retrieve is located

`-i, --id [id]`
: ID of the list to retrieve information for. Specify either `id` or `title` but not both

`-t, --title [title]`
: Title of the list to retrieve information for. Specify either `id` or `title` but not both

`-p, --properties [properties]`
: Comma-separated list of properties to retrieve from the list. Will retrieve all properties possible from default response, if not specified.

`--withPermissions`
: Set if you want to return associated roles and permissions of the list.

--8<-- "docs/cmd/_global.md"

## Examples

Return information about a list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list get --id 0cd891ef-afce-4e55-b836-fce03286cccf --webUrl https://contoso.sharepoint.com/sites/project-x
```

Return information about a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get information about a list returning the specified list properties

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Id,HasUniqueRoleAssignments,AllowContentTypes"
```

Get information about a list along with the roles and permissions

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x --withPermissions
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)

## Response

=== "JSON"

    ```json
    {
      "AllowContentTypes": true,
      "BaseTemplate": 100,
      "BaseType": 0,
      "ContentTypesEnabled": true,
      "CrawlNonDefaultViews": false,
      "Created": "2022-10-23T09:30:00Z",
      "CurrentChangeToken": {
        "StringValue": "1;3;97d19285-b8a6-4c7f-9c6c-d6b850a6561a;638042258222730000;564169620"
      },
      "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000",
      "DefaultItemOpenUseListSetting": false,
      "Description": "",
      "Direction": "none",
      "DisableCommenting": false,
      "DisableGridEditing": false,
      "DocumentTemplateUrl": null,
      "DraftVersionVisibility": 0,
      "EnableAttachments": true,
      "EnableFolderCreation": false,
      "EnableMinorVersions": false,
      "EnableModeration": false,
      "EnableRequestSignOff": true,
      "EnableVersioning": true,
      "EntityTypeName": "TestList",
      "ExemptFromBlockDownloadOfNonViewableFiles": false,
      "FileSavePostProcessingEnabled": false,
      "ForceCheckout": false,
      "HasExternalDataSource": false,
      "Hidden": false,
      "Id": "97d19285-b8a6-4c7f-9c6c-d6b850a6561a",
      "ImagePath": {
        "DecodedUrl": "/_layouts/15/images/itgen.png?rev=47"
      },
      "ImageUrl": "/_layouts/15/images/itgen.png?rev=47",
      "DefaultSensitivityLabelForLibrary": "",
      "IrmEnabled": false,
      "IrmExpire": false,
      "IrmReject": false,
      "IsApplicationList": false,
      "IsCatalog": false,
      "IsPrivate": false,
      "ItemCount": 0,
      "LastItemDeletedDate": "2022-11-16T19:55:37Z",
      "LastItemModifiedDate": "2022-11-16T19:55:39Z",
      "LastItemUserModifiedDate": "2022-11-16T19:55:37Z",
      "ListExperienceOptions": 0,
      "ListItemEntityTypeFullName": "SP.Data.TestListItem",
      "MajorVersionLimit": 50,
      "MajorWithMinorVersionsLimit": 0,
      "MultipleDataList": false,
      "NoCrawl": false,
      "ParentWebPath": {
        "DecodedUrl": "/"
      },
      "ParentWebUrl": "/",
      "ParserDisabled": false,
      "ServerTemplateCanCreateFolders": true,
      "TemplateFeatureId": "00bfea71-de22-43b2-a848-c05709900100",
      "Title": "Test"
    }
    ```

=== "Text"

    ```text
    AllowContentTypes                        : true
    BaseTemplate                             : 100
    BaseType                                 : 0
    ContentTypesEnabled                      : true
    CrawlNonDefaultViews                     : false
    Created                                  : 2022-10-23T09:30:00Z
    CurrentChangeToken                       : {"StringValue":"1;3;97d19285-b8a6-4c7f-9c6c-d6b850a6561a;638042258464070000;564169707"}
    DefaultContentApprovalWorkflowId         : 00000000-0000-0000-0000-000000000000
    DefaultItemOpenUseListSetting            : false
    DefaultSensitivityLabelForLibrary        :
    Description                              :
    Direction                                : none
    DisableCommenting                        : false
    DisableGridEditing                       : false
    DocumentTemplateUrl                      : null
    DraftVersionVisibility                   : 0
    EnableAttachments                        : true
    EnableFolderCreation                     : false
    EnableMinorVersions                      : false
    EnableModeration                         : false
    EnableRequestSignOff                     : true
    EnableVersioning                         : true
    EntityTypeName                           : TestList
    ExemptFromBlockDownloadOfNonViewableFiles: false
    FileSavePostProcessingEnabled            : false
    ForceCheckout                            : false
    HasExternalDataSource                    : false
    Hidden                                   : false
    Id                                       : 97d19285-b8a6-4c7f-9c6c-d6b850a6561a
    ImagePath                                : {"DecodedUrl":"/_layouts/15/images/itgen.png?rev=47"}
    ImageUrl                                 : /_layouts/15/images/itgen.png?rev=47
    IrmEnabled                               : false
    IrmExpire                                : false
    IrmReject                                : false
    IsApplicationList                        : false
    IsCatalog                                : false
    IsPrivate                                : false
    ItemCount                                : 0
    LastItemDeletedDate                      : 2022-11-16T19:55:37Z
    LastItemModifiedDate                     : 2022-11-16T19:55:39Z
    LastItemUserModifiedDate                 : 2022-11-16T19:55:37Z
    ListExperienceOptions                    : 0
    ListItemEntityTypeFullName               : SP.Data.TestListItem
    MajorVersionLimit                        : 50
    MajorWithMinorVersionsLimit              : 0
    MultipleDataList                         : false
    NoCrawl                                  : false
    ParentWebPath                            : {"DecodedUrl":"/"}
    ParentWebUrl                             : /
    ParserDisabled                           : false
    ServerTemplateCanCreateFolders           : true
    TemplateFeatureId                        : 00bfea71-de22-43b2-a848-c05709900100
    Title                                    : Test
    ```

=== "CSV"

    ```csv
    AllowContentTypes,BaseTemplate,BaseType,ContentTypesEnabled,CrawlNonDefaultViews,Created,CurrentChangeToken,DefaultContentApprovalWorkflowId,DefaultItemOpenUseListSetting,Description,Direction,DisableCommenting,DisableGridEditing,DocumentTemplateUrl,DraftVersionVisibility,EnableAttachments,EnableFolderCreation,EnableMinorVersions,EnableModeration,EnableRequestSignOff,EnableVersioning,EntityTypeName,ExemptFromBlockDownloadOfNonViewableFiles,FileSavePostProcessingEnabled,ForceCheckout,HasExternalDataSource,Hidden,Id,ImagePath,ImageUrl,DefaultSensitivityLabelForLibrary,IrmEnabled,IrmExpire,IrmReject,IsApplicationList,IsCatalog,IsPrivate,ItemCount,LastItemDeletedDate,LastItemModifiedDate,LastItemUserModifiedDate,ListExperienceOptions,ListItemEntityTypeFullName,MajorVersionLimit,MajorWithMinorVersionsLimit,MultipleDataList,NoCrawl,ParentWebPath,ParentWebUrl,ParserDisabled,ServerTemplateCanCreateFolders,TemplateFeatureId,Title
    1,100,0,1,,2022-10-23T09:30:00Z,"{""StringValue"":""1;3;97d19285-b8a6-4c7f-9c6c-d6b850a6561a;638042258543870000;564169743""}",00000000-0000-0000-0000-000000000000,,,none,,,,0,1,,,,1,1,TestList,,,,,,97d19285-b8a6-4c7f-9c6c-d6b850a6561a,"{""DecodedUrl"":""/_layouts/15/images/itgen.png?rev=47""}",/_layouts/15/images/itgen.png?rev=47,,,,,,,,0,2022-11-16T19:55:37Z,2022-11-16T19:55:39Z,2022-11-16T19:55:37Z,0,SP.Data.TestListItem,50,0,,,"{""DecodedUrl"":""/""}",/,,1,00bfea71-de22-43b2-a848-c05709900100,Test
    ```
