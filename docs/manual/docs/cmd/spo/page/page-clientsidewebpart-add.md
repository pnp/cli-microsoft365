# spo page clientsidewebpart add

Adds a client-side web part to a modern page

## Usage

```sh
spo page clientsidewebpart add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the page to add the web part to is located
`-n, --pageName <pageName>`|Name of the page to which add the web part
`--standardWebPart [standardWebPart]`|Name of the standard web part to add (see the possible values below)
`--webPartId [webPartId]`|ID of the custom web part to add
`--webPartProperties [webPartProperties]`|JSON string with web part properties to set on the web part
`--section [section]`|Number of the section to which the web part should be added (1 or higher)
`--column [column]`|Number of the column in which the web part should be added (1 or higher)
`--order [order]`|Order of the web part in the column
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To add a client-side web part to a modern page, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified `pageName` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

To add a standard web part to the page, specify one of the following values: _ContentRollup, BingMap, ContentEmbed, DocumentEmbed, Image, ImageGallery, LinkPreview, NewsFeed, NewsReel, PowerBIReportEmbed, QuickChart, SiteActivity, VideoEmbed, YammerEmbed, Events, GroupCalendar, Hero, List, PageTitle, People, QuickLinks, CustomMessageRegion, Divider, MicrosoftForms, Spacer_.

When specifying the JSON string with web part properties on Windows, you have to escape double quotes in a specific way. Considering the following value for the _webPartProperties_ option: `{"Foo":"Bar"}`, you should specify the value as \`"{""Foo"":""Bar""}"\`. In addition, when using PowerShell, you should use the `--%` argument.

## Examples

Add the standard Bing Map web part to a modern page in the first available location on the page

```sh
spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap
```

Add the standard Bing Map web part to a modern page in the third column of the second section

```sh
spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --section 2 --column 3
```

Add a custom web part to the modern page

```sh
spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --webPartId 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1
```

Using PowerShell, add the standard Bing Map web part with the specific properties to a modern page

```PowerShell
o365 --% spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --webPartProperties `"{""Title"":""Foo location""}"`
```

Using Windows command line, add the standard Bing Map web part with the specific properties to a modern page

```sh
o365 spo page clientsidewebpart add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --webPartProperties `"{""Title"":""Foo location""}"`
```