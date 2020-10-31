# spo page add

Creates modern page

## Usage

```sh
m365 spo page add [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: Name of the page to create

`-u, --webUrl <webUrl>`
: URL of the site where the page should be created

`-t, --title [title]`
: Title of the page to create. If not specified, will use the page name as its title

`-l, --layoutType [layoutType]`
: Layout of the page. Allowed values `Article,Home`. Default `Article`

`-p, --promoteAs [promoteAs]`
: Create the page for a specific purpose. Allowed values `HomePage,NewsPage`

`--commentsEnabled`
: Set to enable comments on the page

`--publish`
: Set to publish the page

`--publishMessage [publishMessage]`
: Message to set when publishing the page

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If you try to create a page with a name of a page that already exists, you will get a `The file exists` error.

If you choose to promote the page using the `promoteAs` option or enable page comments, you will see the result only after publishing the page.

## Examples

Create new modern page. Use the Article layout

```sh
m365 spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team
```

Create new modern page and set its title

```sh
m365 spo page add --name new-page.aspx --title 'My page' --webUrl https://contoso.sharepoint.com/sites/a-team
```

Create new modern page. Use the Home page layout and include the default set of web parts

```sh
m365 spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --layoutType Home
```

Create new article page and promote it as a news article

```sh
m365 spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --promoteAs NewsPage
```

Create new page and set it as the site's home page

```sh
m365 spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --layoutType Home --promoteAs HomePage
```

Create new article page and promote it as a template

```sh
m365 spo page add --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --promoteAs Template
```

Create new article page and enable comments on the page

```sh
m365 spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --commentsEnabled
```

Create new article page and publish it

```sh
m365 spo page add --name new-page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --publish
```