# spo hubsite register

Registers the specified site collection as a hub site

## Usage

```sh
m365 spo hubsite register [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site collection to register as a hub site

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global or SharePoint administrator.

If the specified site collection is already registered as a hub site, you will get a `This site is already a HubSite.` error.

## Examples

Register the site collection with URL _https://contoso.sharepoint.com/sites/sales_ as a hub site

```sh
m365 spo hubsite register --siteUrl https://contoso.sharepoint.com/sites/sales
```

## Response

=== "JSON"

    ```json
    {
      "Description": "",
      "EnablePermissionsSync": false,
      "EnforcedECTs": null,
      "EnforcedECTsVersion": 0,
      "HideNameInNavigation": false,
      "ID": "0f9b8f4f-0e8e-4630-bb0a-501442db9b64",
      "LogoUrl": null,
      "ParentHubSiteId": "00000000-0000-0000-0000-000000000000",
      "PermissionsSyncTag": 0,
      "RequiresJoinApproval": false,
      "SiteDesignId": "00000000-0000-0000-0000-000000000000",
      "SiteId": "0f9b8f4f-0e8e-4630-bb0a-501442db9b64",
      "SiteUrl": "https://contoso.sharepoint.com/sites/newHubSite",
      "Targets": null,
      "TenantInstanceId": "4d128b52-7228-46b5-8765-5b338476054d",
      "Title": "New Hub Site"
    }
    ```

=== "Text"

    ```text
    Description          : null
    EnablePermissionsSync: false
    EnforcedECTs         : null
    EnforcedECTsVersion  : 0
    HideNameInNavigation : false
    ID                   : 0f9b8f4f-0e8e-4630-bb0a-501442db9b64
    LogoUrl              : null
    ParentHubSiteId      : 00000000-0000-0000-0000-000000000000
    PermissionsSyncTag   : 0
    RequiresJoinApproval : false
    SiteDesignId         : 00000000-0000-0000-0000-000000000000
    SiteId               : 0f9b8f4f-0e8e-4630-bb0a-501442db9b64
    SiteUrl              : https://contoso.sharepoint.com/sites/newHubsite
    Targets              : null
    TenantInstanceId     : 4d128b52-7228-46b5-8765-5b338476054d
    Title                : New Hub Site
    ```

=== "CSV"

    ```csv
    Description,EnablePermissionsSync,EnforcedECTs,EnforcedECTsVersion,HideNameInNavigation,ID,LogoUrl,ParentHubSiteId,PermissionsSyncTag,RequiresJoinApproval,SiteDesignId,SiteId,SiteUrl,Targets,TenantInstanceId,Title
    ,,,0,,0f9b8f4f-0e8e-4630-bb0a-501442db9b64,,00000000-0000-0000-0000-000000000000,0,,00000000-0000-0000-0000-000000000000,0f9b8f4f-0e8e-4630-bb0a-501442db9b64,https://contoso.sharepoint.com/sites/newHubSite,,4d128b52-7228-46b5-8765-5b338476054d,New Hub Site
    ```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
