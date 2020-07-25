# spo page set

Updates modern page properties

## Usage

```sh
m365 spo page set [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: Name of the page to update

`-u, --webUrl <webUrl>`
: URL of the site where the page to update is located

`-l, --layoutType [layoutType]`
: Layout of the page. Allowed values `Article,Home`

`-p, --promoteAs [promoteAs]`
: Update the page purpose. Allowed values `HomePage,NewsPage`

`--commentsEnabled [commentsEnabled]`
: Set to `true`, to enable comments on the page. Allowed values `true,false`

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

If you try to create a page with a name of a page that already exists, you will get a `The file doesn't exists` error.

If you choose to promote the page using the `promoteAs` option or enable page comments, you will see the result only after publishing the page.

## Examples

Change the layout of the existing page to _Article_

```sh
m365 spo page set --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --layoutType Article
```

Promote the existing article page as a news article

```sh
m365 spo page set --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --promoteAs NewsPage
```

Promote the existing article page as a template

```sh
m365 spo page set --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --promoteAs Template
```

Change the page's layout to Home and set it as the site's home page

```sh
m365 spo page set --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --layoutType Home --promoteAs HomePage
```

Enable comments on the existing page

```sh
m365 spo page set --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --commentsEnabled true
```

Publish existing page

```sh
m365 spo page set --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --publish
```