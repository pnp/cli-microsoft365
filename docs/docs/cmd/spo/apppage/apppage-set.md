# spo apppage set

Updates the single-part app page

## Usage

```sh
m365 spo apppage set [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the page to update is located

`-n, --pageName <pageName>`
: The name of the page to be updated, eg. page.aspx

`-d, --webPartData <webPartData>`
: JSON string of the web part to update on the page

--8<-- "docs/cmd/_global.md"

## Examples

Updates the single-part app page located in a site with url https://contoso.sharepoint.com. Web part data is stored in the `$webPartData` variable

```sh
m365 spo apppage set --webUrl "https://contoso.sharepoint.com" --pageName "Contoso.aspx" --webPartData $webPartData
```
