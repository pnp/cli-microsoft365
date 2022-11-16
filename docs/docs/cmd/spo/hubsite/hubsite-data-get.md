# spo hubsite data get

Get hub site data for the specified site

## Usage

```sh
m365 spo hubsite data get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve hub site data

`-f, --forceRefresh`
: Set, to refresh the server cache with the latest updates

--8<-- "docs/cmd/_global.md"

## Remarks

By default, the hub site data is returned from the server's cache. To refresh the data with the latest updates, use the `-f, --forceRefresh` option. Use this option, if you just made changes and need to see them right
away.

If the specified site is not connected to a hub site site and is not a hub site itself, no data will be retrieved.

## Examples

Get information about the hub site data for a site with URL https://contoso.sharepoint.com/sites/project-x

```sh
m365 spo hubsite data get --webUrl https://contoso.sharepoint.com/sites/project-x
```

## Response

=== "JSON"

    ```json
    {
      "headerEmphasis": "None",
      "themeKey": "7EDE94FF",
      "name": "Intranet",
      "url": "https://contoso.sharepoint.com/sites/intra",
      "logoUrl": "https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png",
      "logoFileHash": 637696294610000000,
      "usesMetadataNavigation": false,
      "megaMenuEnabled": true,
      "navigation": [],
      "isNavAudienceTargeted": false,
      "siteDesignId": "184644fb-90ed-4841-a7ad-6930cf819060",
      "requiresJoinApproval": false,
      "hideNameInNavigation": false,
      "parentHubSiteId": "1e1232eb-1a78-4726-8bb9-56af3640228d",
      "relatedHubSiteIds": [
        "af80c11f-0138-4d72-bb37-514542c3aabb"
      ],
      "tenantInstanceId": "4d128b52-7228-46b5-8765-5b338476054d",
      "isSameTenantInstance": true
    }
    ```

=== "Text"

    ```text
    headerEmphasis        : None
    hideNameInNavigation  : false
    isNavAudienceTargeted : false
    isSameTenantInstance  : true
    logoFileHash          : 637696294610000000
    logoUrl               : https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png
    megaMenuEnabled       : true
    name                  : Intranet
    navigation            : []
    parentHubSiteId       : 1e1232eb-1a78-4726-8bb9-56af3640228d
    relatedHubSiteIds     : ["af80c11f-0138-4d72-bb37-514542c3aabb"]
    requiresJoinApproval  : false
    siteDesignId          : 184644fb-90ed-4841-a7ad-6930cf819060
    tenantInstanceId      : 4d128b52-7228-46b5-8765-5b338476054d
    themeKey              : 7EDE94FF
    url                   : https://contoso.sharepoint.com/sites/intra
    usesMetadataNavigation: false
    ```

=== "CSV"

    ```csv
    headerEmphasis,themeKey,name,url,logoUrl,logoFileHash,usesMetadataNavigation,megaMenuEnabled,navigation,isNavAudienceTargeted,siteDesignId,requiresJoinApproval,hideNameInNavigation,parentHubSiteId,relatedHubSiteIds,tenantInstanceId,isSameTenantInstance
    None,7EDE94FF,Intranet,https://contoso.sharepoint.com/sites/intra,https://contoso.sharepoint.com/sites/intra/SiteAssets/work.png,637696294610000000,false,1,[],false,184644fb-90ed-4841-a7ad-6930cf819060,,,1e1232eb-1a78-4726-8bb9-56af3640228d,"[""af80c11f-0138-4d72-bb37-514542c3aabb""]",4d128b52-7228-46b5-8765-5b338476054d,1
    ```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
