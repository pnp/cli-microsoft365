# spo page copy

Creates a copy of a modern page or template

## Usage

```sh
m365 spo page copy [options]
```

## Options

`--sourceName <sourceName>`
: The name of the source file.

`--targetUrl <targetUrl>`
: The URL of the target file. You can specify page's name or relative- or absolute URL.

`--overwrite`
: Overwrite the target page when it already exists.

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located.

--8<-- "docs/cmd/_global.md"

## Remarks

If another page exists at the specified target location, copying the page will fail, unless you include the `--overwrite` option.

## Examples

Create a copy of the `home.aspx` page

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "home.aspx" --targetUrl "home-copy.aspx"
```

Overwrite the target page if it already exists

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "home.aspx" --targetUrl "home-copy.aspx" --overwrite
```

Create a copy of a page template

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "templates/PageTemplate.aspx" --targetUrl "page.aspx"
```

Copy the page to another site

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "templates/PageTemplate.aspx" --targetUrl "https://contoso.sharepoint.com/sites/team-b/sitepages/page.aspx"
```


## Response

=== "JSON"

    ```json
    {
      "AbsoluteUrl": "https://contoso.sharepoint.com/sites/SPDemo/SitePages/home-copy.aspx",
      "AuthorByline": null,
      "BannerImageUrl": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png",
      "BannerThumbnailUrl": "https://media.akamai.odsp.cdn.office.net/contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png",
      "CallToAction": "",
      "Categories": null,
      "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C411800F1678937A82C3142BEF3C962300813B5",
      "Description": null,
      "DoesUserHaveEditPermission": true,
      "FileName": "home-copy.aspx",
      "FirstPublished": "0001-01-01T08:00:00Z",
      "Id": 23,
      "IsPageCheckedOutToCurrentUser": false,
      "IsWebWelcomePage": false,
      "Modified": "2022-11-26T10:11:30Z",
      "PageLayoutType": "Article",
      "Path": {
        "DecodedUrl": "SitePages/home-copy.aspx"
      },
      "PromotedState": 0,
      "Title": "new-page",
      "TopicHeader": null,
      "UniqueId": "3c4b010b-7043-4b2b-b7eb-e6110d1bebac",
      "Url": "SitePages/home-copy.aspx",
      "Version": "0.1",
      "VersionInfo": {
        "LastVersionCreated": "0001-01-01T00:00:00",
        "LastVersionCreatedBy": ""
      },
      "AlternativeUrlMap": "{\"MediaTAThumbnailPathUrl\":\"https://southindia1-mediap.svc.ms/transform/thumbnail?provider=spo&inputFormat={.fileType}&cs=UEFHRVN8U1BP&docid={.spHost}/_api/v2.0/sharePoint:{.resourceUrl}:/driveItem&w={.widthValue}&oauth_token=bearer%20{.oauthToken}\",\"MediaTAThumbnailHostUrl\":\"https://southindia1-mediap.svc.ms\",\"AFDCDNEnabled\":\"ClientNotOnEdge\",\"CurrentSiteCDNPolicy\":\"True\",\"PublicCDNEnabled\":\"True\",\"PrivateCDNEnabled\":\"True\"}",
      "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"7558d804-0334-49ca-b14a-53870cf6caae\",\"position\":{\"controlIndex\":1,\"sectionIndex\":1,\"zoneIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"webPartId\":\"e377ea37-9047-43b9-8cdb-a761be2f8e09\",\"emphasis\":{},\"webPartData\":{\"dataVersion\":\"1.0\",\"description\":\"Display a location on a map using Bing Maps.\",\"id\":\"e377ea37-9047-43b9-8cdb-a761be2f8e09\",\"instanceId\":\"7558d804-0334-49ca-b14a-53870cf6caae\",\"properties\":{\"pushPins\":[],\"maxNumberOfPushPins\":1,\"shouldShowPushPinTitle\":true,\"zoomLevel\":12,\"mapType\":\"road\"},\"title\":\"Bing Maps\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}}}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]",
      "CoAuthState": null,
      "Language": "en-us",
      "LayoutWebpartsContent": null,
      "SitePageFlags": ""
    }
    ```

=== "Text"

    ```text
    Id            : 24
    PageLayoutType: Article
    Title         : new-page
    Url           : SitePages/home-copy.aspx
    ```

=== "CSV"

    ```csv
    Id,PageLayoutType,Title,Url
    25,Article,new-page,SitePages/home-copy.aspx
    ```
