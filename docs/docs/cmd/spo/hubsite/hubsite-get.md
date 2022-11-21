# spo hubsite get

Gets information about the specified hub site

## Usage

```sh
m365 spo hubsite get [options]
```

## Options

`-i, --id [id]`
: ID of the hub site. Specify either `id`, `title` or `url` but not multiple.

`-t, --title [title]`
: Title of the hub site. Specify either `id`, `title` or `url` but not multiple.

`-u, --url [url]`
: URL of the hub site. Specify either `id`, `title` or `url` but not multiple.

`--includeAssociatedSites`
: Include the associated sites in the result (only in JSON output)

--8<-- "docs/cmd/_global.md"

## Examples

Get information about the hub site with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
m365 spo hubsite get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

Get information about the hub site with Title _My Hub Site_

```sh
m365 spo hubsite get --title 'My Hub Site'
```

Get information about the hub site with URL _https://contoso.sharepoint.com/sites/HubSite_

```sh
m365 spo hubsite get --url 'https://contoso.sharepoint.com/sites/HubSite'
```

Get information about the hub site with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_, including its associated sites. Associated site info is only shown in JSON output.

```sh
m365 spo hubsite get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --includeAssociatedSites --output json
```

Get information about the hub site with Title _My Hub Site_

```sh
m365 spo hubsite get --title "My Hub Site"
```

Get information about the hub site with URL _https://contoso.sharepoint.com/sites/HubSite_

```sh
m365 spo hubsite get --url "https://contoso.sharepoint.com/sites/HubSite"
```

## Response

### Standard response

=== "JSON"

    ```json
    {
      "Description": "Intranet Hub Site",
      "EnablePermissionsSync": false,
      "EnforcedECTs": null,
      "EnforcedECTsVersion": 0,
      "HideNameInNavigation": false,
      "ID": "af80c11f-0138-4d72-bb37-514542c3aabb",
      "LogoUrl": "https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png",
      "ParentHubSiteId": "ec78f3aa-5a74-4f16-be49-3396df045f34",
      "PermissionsSyncTag": 0,
      "RequiresJoinApproval": false,
      "SiteDesignId": "184644fb-90ed-4841-a7ad-6930cf819060",
      "SiteId": "af80c11f-0138-4d72-bb37-514542c3aabb",
      "SiteUrl": "https://contoso.sharepoint.com/sites/intra",
      "Targets": null,
      "TenantInstanceId": "5d128b52-7228-46b5-8765-5b338476054d",
      "Title": "Intranet"
    }
    ```

=== "Text"

    ```text
    Description          : Intranet Hub Site
    EnablePermissionsSync: false
    EnforcedECTs         : null
    EnforcedECTsVersion  : 0
    HideNameInNavigation : false
    ID                   : af80c11f-0138-4d72-bb37-514542c3aabb
    LogoUrl              : https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png
    ParentHubSiteId      : ec78f3aa-5a74-4f16-be49-3396df045f34
    PermissionsSyncTag   : 0
    RequiresJoinApproval : false
    SiteDesignId         : 184644fb-90ed-4841-a7ad-6930cf819060
    SiteId               : af80c11f-0138-4d72-bb37-514542c3aabb
    SiteUrl              : https://contoso.sharepoint.com/sites/intra
    Targets              : null
    TenantInstanceId     : 5d128b52-7228-46b5-8765-5b338476054d
    Title                : Intranet
    ```

=== "CSV"

    ```csv
    Description,EnablePermissionsSync,EnforcedECTs,EnforcedECTsVersion,HideNameInNavigation,ID,LogoUrl,ParentHubSiteId,PermissionsSyncTag,RequiresJoinApproval,SiteDesignId,SiteId,SiteUrl,Targets,TenantInstanceId,Title
    Intranet Hub Site,false,,0,false,af80c11f-0138-4d72-bb37-514542c3aabb,https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png,ec78f3aa-5a74-4f16-be49-3396df045f34,0,false,184644fb-90ed-4841-a7ad-6930cf819060,af80c11f-0138-4d72-bb37-514542c3aabb,https://contoso.sharepoint.com/sites/intra,,5d128b52-7228-46b5-8765-5b338476054d,Intranet
    ```

### `includeAssociatedSites` response

When we make use of the option `includeAssociatedSites` the response will differ. This command can only be executed using --output json or an error will be thrown.

=== "JSON"

    ```json
    {
      "Description": "Intranet Hub Site",
      "EnablePermissionsSync": false,
      "EnforcedECTs": null,
      "EnforcedECTsVersion": 0,
      "HideNameInNavigation": false,
      "ID": "af80c11f-0138-4d72-bb37-514542c3aabb",
      "LogoUrl": "https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png",
      "ParentHubSiteId": "ec78f3aa-5a74-4f16-be49-3396df045f34",
      "PermissionsSyncTag": 0,
      "RequiresJoinApproval": false,
      "SiteDesignId": "184644fb-90ed-4841-a7ad-6930cf819060",
      "SiteId": "af80c11f-0138-4d72-bb37-514542c3aabb",
      "SiteUrl": "https://contoso.sharepoint.com/sites/intra",
      "Targets": null,
      "TenantInstanceId": "5d128b52-7228-46b5-8765-5b338476054d",
      "Title": "Intranet",
      "AssociatedSites": [
        {
          "Title": "About Us",
          "SiteId": "1e1232eb-1a78-4726-8bb9-56af3640228d",
          "SiteUrl": "https://contoso.sharepoint.com/sites/about-us"
        }
      ]
    }
    ```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
