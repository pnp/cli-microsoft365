# spo apppage set

Updates the single-part app page

## Usage

```sh
m365 spo apppage set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the page to update is located

`-n, --pageName <pageName>`
: The name of the page to be updated, eg. page.aspx

`-d, --webPartData <webPartData>`
: JSON string of the web part to update on the page

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Updates the single-part app page located in a site with url https://contoso.sharepoint.com. Web part data is stored in the `$webPartData` variable

```sh
m365 spo apppage set --webUrl "https://contoso.sharepoint.com" --pageName "Contoso.aspx" --webPartData $webPartData
```
