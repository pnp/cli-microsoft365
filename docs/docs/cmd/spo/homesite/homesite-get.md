# spo homesite get

Gets information about the Home Site

## Usage

```sh
m365 spo homesite get [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Get information about the Home Site

```sh
m365 spo homesite get
```

## Response

=== "JSON"

    ```json
    {
      "SiteId": "af80c11f-0138-4d72-bb37-514542c3aabb",
      "WebId": "6f90666d-b0e7-40c3-991f-4ab051d00a70",
      "LogoUrl": "https://contoso.sharepoint.com/sites/intra/siteassets/work.png",
      "Title": "Intranet",
      "Url": "https://contoso.sharepoint.com/sites/intra"
    }
    ```

=== "Text"

    ```text
    LogoUrl: https://contoso.sharepoint.com/sites/intra/siteassets/work.png
    SiteId : af80c11f-0138-4d72-bb37-514542c3aabb
    Title  : Intranet
    Url    : https://contoso.sharepoint.com/sites/intra
    WebId  : 6f90666d-b0e7-40c3-991f-4ab051d00a70
    ```

=== "CSV"

    ```csv
    SiteId,WebId,LogoUrl,Title,Url
    af80c11f-0138-4d72-bb37-514542c3aabb,6f90666d-b0e7-40c3-991f-4ab051d00a70,https://contoso.sharepoint.com/sites/intra/siteassets/work.png,Intranet,https://contoso.sharepoint.com/sites/intra
    ```

## More information

- SharePoint home sites: a landing for your organization on the intelligent intranet: [https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933](https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933)
