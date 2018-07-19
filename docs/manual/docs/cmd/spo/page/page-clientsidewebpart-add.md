# spo page clientsidewebpart add

Adds a Client Side WebPart to a modern page

## Usage

```sh
spo page clientsidewebpart add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --pageName <name>`|Name of the page to add the WebPart to
`-u, --webUrl <webUrl>`|URL of the site where the page to add the WebPart to is located
`--webPartId`|The ID of the WebPart to add to the page
`--standardWebPart`|The identifier of a standard WebPart to add to the page. `ContentRollup|BingMap|ContentEmbed|DocumentEmbed|Image|ImageGallery|LinkPreview|NewsFeed|NewsReel|PowerBIReportEmbed|QuickChart|SiteActivity|VideoEmbed|YammerEmbed|Events|GroupCalendar|Hero|List|PageTitle|People|QuickLinks|CustomMessageRegion|Divider|MicrosoftForms|Spacer`
`--webPartProperties`|The JSON string representing the properties of the WebPart
`--section`|The number of the section to add the WebPart to. First section has number 1
`--column`|The number of the column to add the WebPart to. First column has number 1
`--order`|The order index of the component. Used when multiple WebParts are added to the same section and column
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To create new modern page, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If you specify a section or column that does not exists in the specified Page, you will get a `Invalid Section` or `Invalid Column` error.

### On Windows in non-immersive mode

When using with Windows shells such as PowerShell or CMD, you have to escape double quotes in a specific way in the JSON of the --webPartProperties parameter.
Considering the following value for the --webPartProperties argument: ```{"Foo":"Bar"}``` , should you specify the value as \`"{""Foo"":""Bar""}"\`
In addition, using PowerShell interface, should you use the `--%` argument

e.g. In PowerShell

```PowerShell
o365 --% spo page clientsidewebparts add --pageName page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --standardWebPart BingMap --webPartProperties `"{""Title"":""Foo location""}"`
```

e.g. In CMD

```CMD
o365 spo page clientsidewebparts add --pageName page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --standardWebPart BingMap --webPartProperties `"{""Title"":""Foo location""}"`
```

## Examples

Add a standard WebPart (e.g. Bing Map WebPart) to a page to the first available location 

```sh
spo page clientsidewebparts add --pageName page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --standardWebPart BingMap
```

Add a standard WebPart (e.g. Bing Map WebPart) to a page to second section and third column

```sh
spo page clientsidewebparts add --pageName page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --standardWebPart BingMap --section 2 --column 3
```

Add a standard WebPart (e.g. Bing Map WebPart) to a page to second section and third column at index 2 with specified properties. Properties are stored as JSON string in the `$webPartProps` variable.

```sh
spo page clientsidewebparts add --pageName page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --standardWebPart BingMap --webPartProperties $webPartProps --section 2 --column 3 --order 2
```

Add a WebPart with a specified Id to a page

```sh
spo page clientsidewebparts add --pageName page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --webPartId 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1
```

Add a WebPart with a specified Id to a page with specified properties and specified layout location. Properties are stored as JSON string in the `$webPartProps` variable.

```sh
spo page clientsidewebparts add --pageName page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --webPartId 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webPartProperties $webPartProps --section 1 --column 2 --order 3
```