# spo page text add

Adds text to a modern page

## Usage

```sh
m365 spo page text add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the page to add the text to is located

`-n, --pageName <pageName>`
: Name of the page to which add the text

`-t, --text <text>`
: Text to add to the page

`--section [section]`
: Number of the section to which the text should be added (1 or higher)

`--column [column]`
: Number of the column in which the text should be added (1 or higher)

`--order [order]`
: Order of the text in the column

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified `pageName` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

Add text to a modern page in the first available location on the page

```sh
m365 spo page text add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world'
```

Add text to a modern page in the third column of the second section

```sh
m365 spo page text add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world' --section 2 --column 3
```

Add text at the beginning of the default column on a modern page

```sh
m365 spo page text add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world' --order 1
```