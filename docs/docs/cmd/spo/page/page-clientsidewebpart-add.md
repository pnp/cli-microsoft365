# spo page clientsidewebpart add

Adds a client-side web part to a modern page

## Usage

```sh
m365 spo page clientsidewebpart add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the page to add the web part to is located.

`-n, --pageName <pageName>`
: Name of the page to which add the web part.

`--standardWebPart [standardWebPart]`
: Name of the standard web part to add (see the possible values below). Specify either `standardWebPart` or `webPartId` but not both.

`--webPartId [webPartId]`
: ID of the custom web part to add. Specify either `standardWebPart` or `webPartId` but not both.

`--webPartProperties [webPartProperties]`
: JSON string with web part properties to set on the web part. Specify either `webPartProperties` or `webPartData` but not both.

`--webPartData [webPartData]`
: JSON string with web part data as retrieved from the web part maintenance mode. Specify either `webPartProperties` or `webPartData` but not both.

`--section [section]`
: Number of the section to which the web part should be added (1 or higher).

`--column [column]`
: Number of the column in which the web part should be added (1 or higher).

`--order [order]`
: Order of the web part in the column.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `pageName` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

To add a standard web part to the page, specify one of the following values: _ContentRollup, BingMap, ContentEmbed, DocumentEmbed, Image, ImageGallery, LinkPreview, NewsFeed, NewsReel, PowerBIReportEmbed, QuickChart, SiteActivity, VideoEmbed, YammerEmbed, Events, GroupCalendar, Hero, List, PageTitle, People, QuickLinks, CustomMessageRegion, Divider, MicrosoftForms, Spacer_.

!!! warning "Escaping JSON in PowerShell"
    When using the `--webPartProperties` option it's possible to enter a JSON string. In PowerShell 5 to 7.2 [specific escaping rules](./../../../user-guide/using-cli.md#escaping-double-quotes-in-powershell) apply due to an issue. Remember that you can also use [file tokens](./../../../user-guide/using-cli.md#passing-complex-content-into-cli-options) instead.

## Examples

Add the standard Bing Map web part to a modern page in the first available location on the page

```sh
m365 spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap
```

Add the standard Bing Map web part to a modern page in the third column of the second section

```sh
m365 spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --section 2 --column 3
```

Add a custom web part to the modern page

```sh
m365 spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --webPartId 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1
```

Add the standard Bing Map web part with the specific properties to a modern page

```sh
m365 spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --webPartProperties '{"Title":"Foo location"}'
```

Add the standard Image web part with the preconfigured image

```sh
m365 spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart Image --webPartData '{ "dataVersion": "1.8", "serverProcessedContent": {"htmlStrings":{},"searchablePlainTexts":{"captionText":""},"imageSources":{"imageSource":"/sites/team-a/SiteAssets/work-life-balance.png"},"links":{}}, "properties": {"imageSourceType":2,"altText":"a group of people on a beach","overlayText":"Work life balance","fileName":"48146-OFF12_Justice_01.png","siteId":"27664b85-067d-4be9-a7d7-89b2e804d09f","webId":"a7664b85-067d-4be9-a7d7-89b2e804d09f","listId":"37664b85-067d-4be9-a7d7-89b2e804d09f","uniqueId":"67664b85-067d-4be9-a7d7-89b2e804d09f","imgWidth":650,"imgHeight":433,"fixAspectRatio":false,"isOverlayTextEnabled":true}}'
```

## Response

The command won't return a response on success.
