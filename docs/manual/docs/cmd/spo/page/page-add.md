# spo page add

Creates modern page

## Usage

```sh
spo page add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the page to create
`-u, --webUrl <webUrl>`|URL of the site where the page should be created
`-l, --layoutType [layoutType]`|Layout of the page. Allowed values `Article|Home`. Default `Article`
`-p, --promoteAs [promoteAs]`|Create the page for a specific purpose. Allowed values `HomePage|NewsPage`
`--commentsEnabled`|Set to enable comments on the page
`--publish`|Set to publish the page
`--publishMessage [publishMessage]`|Message to set when publishing the page
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To create new modern page, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If you try to create a page with a name of a page that already exists, you will get a `The file exists` error.

If you choose to promote the page using the `promoteAs` option or enable page comments, you will see the result only after publishing the page.

## Examples

Create new modern page. Use the Article layout

```sh
spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team
```

Create new modern page. Use the Home page layout and include the default set of web parts

```sh
spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --layoutType Home
```

Create new article page and promote it as a news article

```sh
spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --promoteAs NewsPage
```

Create new page and set it as the site's home page

```sh
spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --layoutType Home --promoteAs HomePage
```

Create new article page and enable comments on the page

```sh
spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --commentsEnabled
```

Create new article page and publish it

```sh
spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --publish
```