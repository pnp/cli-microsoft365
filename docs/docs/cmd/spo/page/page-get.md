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
=== "Markdown"

    ```md
    # spo page get --webUrl "https://contoso.sharepoint.com/sites/team-a" --name "home.aspx"

    Date: 5/1/2023

    ## Home (86ce2216-83f1-4ab4-9b7e-ffbdcf890992)

    Property | Value
    ---------|-------
    CheckInComment | 
    CheckOutType | 2
    ContentTag | {86CE2216-83F1-4AB4-9B7E-FFBDCF890992},14,1
    CustomizedPageStatus | 1
    ETag | "{86CE2216-83F1-4AB4-9B7E-FFBDCF890992},14"
    Exists | true
    IrmEnabled | false
    Length | 805
    Level | 1
    LinkingUrl | 
    MajorVersion | 2
    MinorVersion | 0
    Name | home.aspx
    ServerRelativeUrl | /sites/Company311/SitePages/home.aspx
    TimeCreated | 2021-09-19T00:20:25Z
    TimeLastModified | 2023-05-01T20:43:12Z
    Title | Home
    UIVersion | 1024
    UIVersionLabel | 2.0
    UniqueId | 86ce2216-83f1-4ab4-9b7e-ffbdcf890992
    commentsDisabled | true
    title | Home
    layoutType | Home
    canvasContentJson | [{"position":{"layoutIndex":1,"zoneIndex":0.5,"sectionIndex":1,"controlIndex":1,"sectionFactor":6},"controlType":3,"id":"f85f8dfa-9052-4be8-8954-8cdafe811b97","webPartId":"e377ea37-9047-43b9-8cdb-a761be2f8e09","reservedHeight":528,"reservedWidth":570,"addedFromPersistedData":true,"webPartData":{"id":"e377ea37-9047-43b9-8cdb-a761be2f8e09","instanceId":"f85f8dfa-9052-4be8-8954-8cdafe811b97","title":"Bing Maps","description":"Display a location on a map using Bing Maps.","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.0","properties":{"pushPins":[],"maxNumberOfPushPins":1,"shouldShowPushPinTitle":true,"zoomLevel":12,"mapType":"road","center":{"latitude":51.399299621582038,"longitude":-0.256799995899204,"altitude":0,"altitudeReference":-1}},"containsDynamicDataSource":false}},{"position":{"layoutIndex":1,"zoneIndex":0.5,"sectionIndex":2,"controlIndex":1,"sectionFactor":6},"id":"emptySection","addedFromPersistedData":true},{"controlType":3,"webPartId":"8c88f208-6c77-4bdb-86a0-0c47b4316588","position":{"zoneIndex":1,"sectionIndex":1,"controlIndex":1,"sectionFactor":8},"id":"71eab4c9-8340-4706-96b9-331527890975","addedFromPersistedData":true,"reservedHeight":406,"reservedWidth":776,"webPartData":{"id":"8c88f208-6c77-4bdb-86a0-0c47b4316588","instanceId":"71eab4c9-8340-4706-96b9-331527890975","title":"News","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{"baseUrl":"/sites/Company311"}},"dataVersion":"1.12","properties":{"layoutId":"FeaturedNews","dataProviderId":"news","emptyStateHelpItemsCount":"1","showChrome":true,"carouselSettings":{"autoplay":false,"autoplaySpeed":5,"dots":true,"lazyLoad":true},"showNewsMetadata":{"showSocialActions":false,"showAuthor":true,"showDate":true},"newsDataSourceProp":1,"carouselHeroWrapperComponentId":"","prefetchCount":4,"filters":[{"filterType":1,"value":"","values":[]}],"newsSiteList":[],"renderItemsSliderValue":4,"layoutComponentId":"","webId":"c59dae7c-48bd-4241-96b7-b81d4bbc25cb","siteId":"9fcdddc5-bd7e-4120-b934-bf675b76855f","filterKQLQuery":""},"containsDynamicDataSource":false}},{"position":{"zoneIndex":1,"sectionIndex":1,"controlIndex":1.5,"sectionFactor":8,"layoutIndex":1},"controlType":3,"id":"dcfa91cf-96fa-47cd-9363-7f2fea373526","webPartId":"490d7c76-1824-45b2-9de3-676421c997fa","reservedHeight":326,"reservedWidth":776,"addedFromPersistedData":true,"webPartData":{"id":"490d7c76-1824-45b2-9de3-676421c997fa","instanceId":"dcfa91cf-96fa-47cd-9363-7f2fea373526","title":"Embed","description":"Embed content from other sites such as Sway, YouTube, Vimeo, and more","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.2","properties":{"embedCode":"","cachedEmbedCode":"","shouldScaleWidth":true,"tempState":{},"thumbnailUrl":""},"containsDynamicDataSource":false}},{"controlType":3,"webPartId":"eb95c819-ab8f-4689-bd03-0c2d65d47b1f","position":{"zoneIndex":1,"sectionIndex":1,"controlIndex":2,"sectionFactor":8},"id":"3dde98e4-07d7-46d6-b57f-128d6ae5438c","addedFromPersistedData":true,"reservedHeight":939,"reservedWidth":776,"webPartData":{"id":"eb95c819-ab8f-4689-bd03-0c2d65d47b1f","instanceId":"3dde98e4-07d7-46d6-b57f-128d6ae5438c","title":"Site activity","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.0","properties":{"maxItems":9},"containsDynamicDataSource":false}},{"position":{"zoneIndex":1,"sectionIndex":2,"controlIndex":0.5,"sectionFactor":4,"layoutIndex":1},"controlType":3,"id":"9d1fed5d-9274-4a82-8891-58bbc937c92b","webPartId":"f6fdf4f8-4a24-437b-a127-32e66a5dd9b4","addedFromPersistedData":true,"reservedHeight":449,"reservedWidth":364,"webPartData":{"id":"f6fdf4f8-4a24-437b-a127-32e66a5dd9b4","instanceId":"9d1fed5d-9274-4a82-8891-58bbc937c92b","title":"Twitter","description":"Display a Twitter feed","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.0","properties":{"displayAs":"list","displayHeader":false,"displayFooter":false,"displayBorders":true,"limit":"3","term":"@microsoft","widthSlider":100,"title":"","allowStretch":false,"displayLightTheme":true},"containsDynamicDataSource":false}},{"controlType":3,"webPartId":"c70391ea-0b10-4ee9-b2b4-006d3fcad0cd","position":{"zoneIndex":1,"sectionIndex":2,"controlIndex":1,"sectionFactor":4},"id":"2d412c2a-ed7d-4117-bf54-f1b9a28c1346","addedFromPersistedData":true,"reservedHeight":173,"reservedWidth":364,"webPartData":{"id":"c70391ea-0b10-4ee9-b2b4-006d3fcad0cd","instanceId":"2d412c2a-ed7d-4117-bf54-f1b9a28c1346","title":"Quick links","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{"title":"Quick links","items[0].title":"Learn about a team site","items[1].title":"Learn how to add a page"},"imageSources":{},"links":{"baseUrl":"/sites/Company311","items[0].sourceItem.url":"https://go.microsoft.com/fwlink/p/?linkid=827918","items[1].sourceItem.url":"https://go.microsoft.com/fwlink/p/?linkid=827919"},"componentDependencies":{"layoutComponentId":"706e33c8-af37-4e7b-9d22-6e5694d92a6f"}},"dataVersion":"2.2","properties":{"items":[{"sourceItem":{"itemType":2},"thumbnailType":3,"id":1,"description":"","altText":"","rawPreviewImageMinCanvasWidth":32767},{"sourceItem":{"itemType":2},"thumbnailType":3,"id":2,"description":"","altText":"","rawPreviewImageMinCanvasWidth":32767}],"isMigrated":true,"layoutId":"List","shouldShowThumbnail":true,"hideWebPartWhenEmpty":true,"dataProviderId":"QuickLinks","listLayoutOptions":{"showDescription":false,"showIcon":true},"imageWidth":100,"buttonLayoutOptions":{"showDescription":false,"buttonTreatment":2,"iconPositionType":2,"textAlignmentVertical":2,"textAlignmentHorizontal":2,"linesOfText":2},"waffleLayoutOptions":{"iconSize":1,"onlyShowThumbnail":false},"webId":"c59dae7c-48bd-4241-96b7-b81d4bbc25cb","siteId":"9fcdddc5-bd7e-4120-b934-bf675b76855f"},"containsDynamicDataSource":false}},{"controlType":3,"webPartId":"f92bf067-bc19-489e-a556-7fe95f508720","position":{"zoneIndex":1,"sectionIndex":2,"controlIndex":2,"sectionFactor":4},"id":"c047c8d5-b5d0-4852-8a38-b58da702243c","addedFromPersistedData":true,"reservedHeight":291,"reservedWidth":364,"webPartData":{"id":"f92bf067-bc19-489e-a556-7fe95f508720","instanceId":"c047c8d5-b5d0-4852-8a38-b58da702243c","title":"Document library","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{"listTitle":"Documents"},"imageSources":{},"links":{}},"dynamicDataPaths":{},"dynamicDataValues":{"filterBy":{}},"dataVersion":"1.0","properties":{"isDocumentLibrary":true,"showDefaultDocumentLibrary":true,"webpartHeightKey":4,"selectedListUrl":""},"containsDynamicDataSource":true}},{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":true,"isSpellCheckEnabled":true,"globalRichTextStylingVersion":0,"rtePageSettings":{"contentVersion":4},"isEmailReady":false}}]
    numControls | 9
    numSections | 2
    ```
