# spo page get

Gets information about the specific modern page

## Usage

```sh
m365 spo page get [options]
```

## Options

`-n, --name <name>`
: Name of the page to retrieve.

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located.

`--metadataOnly`
: Specify to only retrieve the metadata without the section and control information.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified name doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

## Examples

Get information about the modern page

```sh
m365 spo page get --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```

Get all the metadata from the modern page, without the section and control count information

```sh
m365 spo page get --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx --metadataOnly
```

## Response

=== "JSON"

    ```json
    {
      "ListItemAllFields": {
        "CommentsDisabled": true,
        "FileSystemObjectType": 0,
        "Id": 21,
        "ServerRedirectedEmbedUri": null,
        "ServerRedirectedEmbedUrl": "",
        "ContentTypeId": "0x0101009D1CB255DA76424F860D91F20E6C411800F1678937A82C3142BEF3C962300813B5",
        "OData__ModerationComments": null,
        "ComplianceAssetId": null,
        "WikiField": null,
        "Title": "new-page",
        "ClientSideApplicationId": "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
        "PageLayoutType": "Article",
        "CanvasContent1": "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;controlType&quot;&#58;3,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;7558d804-0334-49ca-b14a-53870cf6caae&quot;,&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;layoutIndex&quot;&#58;1&#125;,&quot;webPartId&quot;&#58;&quot;e377ea37-9047-43b9-8cdb-a761be2f8e09&quot;,&quot;emphasis&quot;&#58;&#123;&#125;&#125;\"><div data-sp-webpart=\"\" data-sp-webpartdataversion=\"1.0\" data-sp-webpartdata=\"&#123;&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;description&quot;&#58;&quot;Display a location on a map using Bing Maps.&quot;,&quot;id&quot;&#58;&quot;e377ea37-9047-43b9-8cdb-a761be2f8e09&quot;,&quot;instanceId&quot;&#58;&quot;7558d804-0334-49ca-b14a-53870cf6caae&quot;,&quot;properties&quot;&#58;&#123;&quot;pushPins&quot;&#58;[],&quot;maxNumberOfPushPins&quot;&#58;1,&quot;shouldShowPushPinTitle&quot;&#58;true,&quot;zoomLevel&quot;&#58;12,&quot;mapType&quot;&#58;&quot;road&quot;&#125;,&quot;title&quot;&#58;&quot;Bing Maps&quot;&#125;\"><div data-sp-componentid=\"e377ea37-9047-43b9-8cdb-a761be2f8e09\"></div><div data-sp-htmlproperties=\"\"></div></div></div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true&#125;&#125;\"></div></div>",
        "BannerImageUrl": {
          "Description": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png",
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
        "OData__ModerationStatus": 3,
        "ID": 21,
        "Created": "2022-11-26T01:51:46",
        "AuthorId": 7,
        "Modified": "2022-11-26T01:55:47",
        "EditorId": 7,
        "OData__CopySource": null,
        "CheckoutUserId": null,
        "OData__UIVersionString": "0.4",
        "GUID": "c8e64e90-e546-4b67-ad05-44e76dac54fb"
      },
      "CheckInComment": "",
      "CheckOutType": 2,
      "ContentTag": "{C431F2EF-447C-4F72-BC3E-ED2687456C33},8,3",
      "CustomizedPageStatus": 2,
      "ETag": "\"{C431F2EF-447C-4F72-BC3E-ED2687456C33},8\"",
      "Exists": true,
      "IrmEnabled": false,
      "Length": "4106",
      "Level": 2,
      "LinkingUri": null,
      "LinkingUrl": "",
      "MajorVersion": 0,
      "MinorVersion": 4,
      "Name": "new-page.aspx",
      "ServerRelativeUrl": "/sites/SPDemo/SitePages/new-page.aspx",
      "TimeCreated": "2022-11-26T09:51:46Z",
      "TimeLastModified": "2022-11-26T09:55:46Z",
      "Title": "new-page",
      "UIVersion": 4,
      "UIVersionLabel": "0.4",
      "UniqueId": "c431f2ef-447c-4f72-bc3e-ed2687456c33",
      "commentsDisabled": true,
      "title": "new-page",
      "layoutType": "Article",
      "canvasContentJson": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"7558d804-0334-49ca-b14a-53870cf6caae\",\"position\":{\"controlIndex\":1,\"sectionIndex\":1,\"zoneIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"webPartId\":\"e377ea37-9047-43b9-8cdb-a761be2f8e09\",\"emphasis\":{},\"webPartData\":{\"dataVersion\":\"1.0\",\"description\":\"Display a location on a map using Bing Maps.\",\"id\":\"e377ea37-9047-43b9-8cdb-a761be2f8e09\",\"instanceId\":\"7558d804-0334-49ca-b14a-53870cf6caae\",\"properties\":{\"pushPins\":[],\"maxNumberOfPushPins\":1,\"shouldShowPushPinTitle\":true,\"zoomLevel\":12,\"mapType\":\"road\"},\"title\":\"Bing Maps\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}}}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]",
      "numControls": 2,
      "numSections": 1
    }
    ```

=== "Text"

    ```text
    commentsDisabled: true
    layoutType      : Article
    numControls     : 2
    numSections     : 1
    title           : new-page
    ```

=== "CSV"

    ```csv
    commentsDisabled,numSections,numControls,title,layoutType
    1,1,2,new-page,Article
    ```
