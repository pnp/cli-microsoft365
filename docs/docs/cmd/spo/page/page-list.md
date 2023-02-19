# spo page list

Lists all modern pages in the given site

## Usage

```sh
m365 spo page list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site from which to retrieve available pages.

--8<-- "docs/cmd/_global.md"

## Examples

List all modern pages in the specific site

```sh
m365 spo page list --webUrl https://contoso.sharepoint.com/sites/team-a
```

## Response

=== "JSON"

    ```json
    [
      {
        "CheckInComment": "",
        "CheckOutType": 2,
        "ContentTag": "{5A5E4B79-C3F7-479A-B9E5-9A462696C92A},1,1",
        "CustomizedPageStatus": 0,
        "ETag": "\"{5A5E4B79-C3F7-479A-B9E5-9A462696C92A},1\"",
        "Exists": true,
        "IrmEnabled": false,
        "Length": 2666,
        "Level": 1,
        "LinkingUri": null,
        "LinkingUrl": "",
        "MajorVersion": 1,
        "MinorVersion": 0,
        "Name": "Home.aspx",
        "ServerRelativeUrl": "/sites/SPDemo/SitePages/Home.aspx",
        "TimeCreated": "2020-09-06T09:18:59Z",
        "TimeLastModified": "2020-09-06T09:18:59Z",
        "Title": "Home",
        "UIVersion": 512,
        "UIVersionLabel": "1.0",
        "UniqueId": "5a5e4b79-c3f7-479a-b9e5-9a462696c92a",
        "ListItemAllFields": {
          "FileSystemObjectType": 0,
          "Id": 10,
          "ServerRedirectedEmbedUri": null,
          "ServerRedirectedEmbedUrl": "",
          "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C411800F1678937A82C3142BEF3C962300813B5",
          "OData__ModerationComments": null,
          "ComplianceAssetId": null,
          "WikiField": null,
          "Title": "Home",
          "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
          "CanvasContent1": "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;controlType&quot;&#58;3,&quot;webPartId&quot;&#58;&quot;e89b5ad5-9ab5-4730-a66b-e1f68994598c&quot;,&quot;id&quot;&#58;&quot;8bba86eb-3174-4917-b2b8-417af1102351&quot;&#125;\"><div data-sp-webpart=\"\" data-sp-webpartdataversion=\"1.0\" data-sp-webpartdata=\"&#123;&quot;id&quot;&#58;&quot;e89b5ad5-9ab5-4730-a66b-e1f68994598c&quot;,&quot;instanceId&quot;&#58;&quot;8bba86eb-3174-4917-b2b8-417af1102351&quot;,&quot;title&quot;&#58;&quot;ReactProvisionAssets&quot;,&quot;description&quot;&#58;&quot;ReactProvisionAssets description&quot;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;description&quot;&#58;&quot;ReactProvisionAssets&quot;&#125;&#125;\"><div data-sp-componentid=\"\">e89b5ad5-9ab5-4730-a66b-e1f68994598c</div><div data-sp-htmlproperties=\"\"></div></div></div></div>",
          "BannerImageUrl": {
            "Description": "/_layouts/15/images/sitepagethumbnail.png",
            "Url": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png"
          },
          "Description": null,
          "PromotedState": 0,
          "FirstPublishedDate": null,
          "LayoutWebpartsContent": null,
          "OData__AuthorBylineId": null,
          "_AuthorBylineStringId": null,
          "OData__TopicHeader": null,
          "OData__SPSitePageFlags": null,
          "OData__OriginalSourceUrl": null,
          "OData__OriginalSourceSiteId": null,
          "OData__OriginalSourceWebId": null,
          "OData__OriginalSourceListId": null,
          "OData__OriginalSourceItemId": null,
          "OData__SPCallToAction": null,
          "OData__ModerationStatus": 0,
          "ID": 10,
          "Created": "2020-09-06T02:18:55-07:00",
          "AuthorId": 1073741823,
          "Modified": "2020-09-06T02:18:55-07:00",
          "EditorId": 1073741823,
          "OData__CopySource": null,
          "CheckoutUserId": null,
          "OData__UIVersionString": "1.0",
          "GUID": "3762c79f-a768-40f0-bc7b-4d2ab03f22e0"
        },
        "AbsoluteUrl": "https://contoso.sharepoint.com/sites/SPDemo/SitePages/Home.aspx",
        "AuthorByline": null,
        "BannerImageUrl": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png",
        "BannerThumbnailUrl": "https://media.akamai.odsp.cdn.office.net/contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png",
        "CallToAction": "",
        "Categories": null,
        "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C411800F1678937A82C3142BEF3C962300813B5",
        "Description": null,
        "DoesUserHaveEditPermission": true,
        "FileName": "Home.aspx",
        "FirstPublished": "0001-01-01T08:00:00Z",
        "Id": 10,
        "IsPageCheckedOutToCurrentUser": false,
        "IsWebWelcomePage": false,
        "Modified": "2020-09-06T09:18:59Z",
        "PageLayoutType": "Article",
        "Path": {
          "DecodedUrl": "SitePages/Home.aspx"
        },
        "PromotedState": 0,
        "TopicHeader": null,
        "Url": "SitePages/Home.aspx",
        "Version": "1.0",
        "VersionInfo": {
          "LastVersionCreated": "0001-01-01T00:00:00-08:00",
          "LastVersionCreatedBy": ""
        },
        "AlternativeUrlMap": "{\"MediaTAThumbnailPathUrl\":\"https://southindia1-mediap.svc.ms/transform/thumbnail?provider=spo&inputFormat={.fileType}&cs=UEFHRVN8U1BP&docid={.spHost}/_api/v2.0/sharePoint:{.resourceUrl}:/driveItem&w={.widthValue}&oauth_token=bearer%20{.oauthToken}\",\"MediaTAThumbnailHostUrl\":\"https://southindia1-mediap.svc.ms\",\"AFDCDNEnabled\":\"ClientNotOnEdge\",\"CurrentSiteCDNPolicy\":\"True\",\"PublicCDNEnabled\":\"True\",\"PrivateCDNEnabled\":\"True\"}",
        "CanvasContent1": "[{\"controlType\":3,\"webPartId\":\"e89b5ad5-9ab5-4730-a66b-e1f68994598c\",\"id\":\"8bba86eb-3174-4917-b2b8-417af1102351\",\"webPartData\":{\"id\":\"e89b5ad5-9ab5-4730-a66b-e1f68994598c\",\"instanceId\":\"8bba86eb-3174-4917-b2b8-417af1102351\",\"title\":\"ReactProvisionAssets\",\"description\":\"ReactProvisionAssets description\",\"dataVersion\":\"1.0\",\"properties\":{\"description\":\"ReactProvisionAssets\"},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}}}}]",
        "CoAuthState": null,
        "Language": "en-us",
        "LayoutWebpartsContent": null,
        "SitePageFlags": ""
      }
    ]
    ```

=== "Text"

    ```text
    Name         Title
    -----------  -------    
    Home.aspx    Home
    ```

=== "CSV"

    ```csv
    Name,Title
    Home.aspx,Home
    ```
