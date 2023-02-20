# spo hubsite list

Lists hub sites in the current tenant

## Usage

```sh
m365 spo hubsite list [options]
```

## Options

`-i, --includeAssociatedSites`
: Include the associated sites in the result (only in JSON output).

--8<-- "docs/cmd/_global.md"

## Remarks

When using the text or csv output type, the command lists only the values of the `ID`, `SiteUrl` and `Title` properties of the hub site. With the output type as JSON, all available properties are included in the command output.

## Examples

List hub sites in the current tenant

```sh
m365 spo hubsite list
```

List hub sites, including their associated sites, in the current tenant. Associated site info is only shown in JSON output.

```sh
m365 spo hubsite list --includeAssociatedSites --output json
```

## Response

### Standard response

=== "JSON"

    ```json
    [
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
    ]
    ```

=== "Text"

    ```text
    ID                                    SiteUrl                                     Title
    ------------------------------------  ------------------------------------------  --------
    af80c11f-0138-4d72-bb37-514542c3aabb  https://contoso.sharepoint.com/sites/intra  Intranet
    ```

=== "CSV"

    ```csv
    ID,SiteUrl,Title
    af80c11f-0138-4d72-bb37-514542c3aabb,https://contoso.sharepoint.com/sites/intra,Intranet
    ```

=== "Markdown"

    ```md
    # spo hubsite list

    Date: 2/20/2023

    ## Intranet (af80c11f-0138-4d72-bb37-514542c3aabb)

    Property | Value
    ---------|-------
    Description | Intranet Hub Site
    EnablePermissionsSync | false
    EnforcedECTs | null
    EnforcedECTsVersion | 0
    HideNameInNavigation | false
    ID | af80c11f-0138-4d72-bb37-514542c3aabb
    LogoUrl | https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png
    ParentHubSiteId | ec78f3aa-5a74-4f16-be49-3396df045f34
    PermissionsSyncTag | 0
    RequiresJoinApproval | false
    SiteDesignId | 184644fb-90ed-4841-a7ad-6930cf819060
    SiteId | af80c11f-0138-4d72-bb37-514542c3aabb
    SiteUrl | https://contoso.sharepoint.com/sites/intra
    Targets | null
    TenantInstanceId | 5d128b52-7228-46b5-8765-5b338476054d
    Title | Intranet
    ```

### `includeAssociatedSites` response

When we make use of the option `includeAssociatedSites` the response will differ. 

=== "JSON"

    ```json
    [
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
            "SiteUrl": "https://contoso.sharepoint.com/sites/about-us"
          }
        ]
      }
    ]
    ```

=== "Text"

    ```text
    ID                                    SiteUrl                                     Title
    ------------------------------------  ------------------------------------------  --------
    af80c11f-0138-4d72-bb37-514542c3aabb  https://contoso.sharepoint.com/sites/intra  Intranet
    ```

=== "CSV"

    ```csv
    ID,SiteUrl,Title
    af80c11f-0138-4d72-bb37-514542c3aabb,https://contoso.sharepoint.com/sites/intra,Intranet
    ```

=== "Markdown"

    ```md
    # spo hubsite list --includeAssociatedSites "true"

    Date: 2/20/2023

    ## Intranet (af80c11f-0138-4d72-bb37-514542c3aabb)

    Property | Value
    ---------|-------
    Description | Intranet Hub Site
    EnablePermissionsSync | false
    EnforcedECTs | null
    EnforcedECTsVersion | 0
    HideNameInNavigation | false
    ID | af80c11f-0138-4d72-bb37-514542c3aabb
    LogoUrl | https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png
    ParentHubSiteId | ec78f3aa-5a74-4f16-be49-3396df045f34
    PermissionsSyncTag | 0
    RequiresJoinApproval | false
    SiteDesignId | 184644fb-90ed-4841-a7ad-6930cf819060
    SiteId | af80c11f-0138-4d72-bb37-514542c3aabb
    SiteUrl | https://contoso.sharepoint.com/sites/intra
    Targets | null
    TenantInstanceId | 5d128b52-7228-46b5-8765-5b338476054d
    Title | Intranet
    ```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
