# spo hubsite set

Updates properties of the specified hub site

## Usage

```sh
m365 spo hubsite set [options]
```

## Options

`-i, --id <id>`
: ID of the hub site to update

`-t, --title [title]`
: The new title for the hub site

`-d, --description [description]`
: The new description for the hub site

`-l, --logoUrl [logoUrl]`
: The URL of the new logo for the hub site

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global or SharePoint administrator.

If the specified `id` doesn't refer to an existing hub site, you will get an `Unknown Error` error.

## Examples

Update hub site's title

```sh
m365 spo hubsite set --id 255a50b2-527f-4413-8485-57f4c17a24d1 --title Sales
```

Update hub site's title and description

```sh
m365 spo hubsite set --id 255a50b2-527f-4413-8485-57f4c17a24d1 --title Sales --description "All things sales"
```

## Response

=== "JSON"

    ```json
    {
      "Description": "Hello",
      "EnablePermissionsSync": false,
      "HideNameInNavigation": false,
      "ID": "af80c11f-0138-4d72-bb37-514542c3aabb",
      "LogoUrl": "https://contoso.sharepoint.com/sites/intra/SiteAssets/teapoint.png",
      "ParentHubSiteId": "/Guid(00000000-0000-0000-0000-000000000000)/",
      "Permissions": null,
      "RequiresJoinApproval": false,
      "SiteDesignId": "/Guid(184644fb-90ed-4841-a7ad-6930cf819060)/",
      "SiteId": "af80c11f-0138-4d72-bb37-514542c3aabb",
      "SiteUrl": "https://contoso.sharepoint.com/sites/intra",
      "Title": "Intranet"
    }
    ```

=== "Text"

    ```text
    Description          : Hello
    EnablePermissionsSync: false
    HideNameInNavigation : false
    ID                   : af80c11f-0138-4d72-bb37-514542c3aabb
    LogoUrl              : https://contoso.sharepoint.com/sites/intra/SiteAssets/teapoint.png
    ParentHubSiteId      : /Guid(00000000-0000-0000-0000-000000000000)/
    Permissions          : null
    RequiresJoinApproval : false
    SiteDesignId         : /Guid(184644fb-90ed-4841-a7ad-6930cf819060)/
    SiteId               : af80c11f-0138-4d72-bb37-514542c3aabb
    SiteUrl              : https://contoso.sharepoint.com/sites/intra
    Title                : Intranet
    ```

=== "CSV"

    ```csv
    Description,EnablePermissionsSync,HideNameInNavigation,ID,LogoUrl,ParentHubSiteId,Permissions,RequiresJoinApproval,SiteDesignId,SiteId,SiteUrl,Title
    Hello,,,af80c11f-0138-4d72-bb37-514542c3aabb,https://contoso.sharepoint.com/sites/intra/SiteAssets/teapoint.png,/Guid(00000000-0000-0000-0000-000000000000)/,,,/Guid(184644fb-90ed-4841-a7ad-6930cf819060)/,af80c11f-0138-4d72-bb37-514542c3aabb,https://contoso.sharepoint.com/sites/intra,Intranet
    ```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
