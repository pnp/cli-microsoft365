# spo apppage add

Creates a single-part app page

## Usage

```sh
m365 spo apppage add [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the page should be created

`-t, --title <title>`
: The title of the page to be created

`-d, --webPartData <webPartData>`
: JSON string of the web part to put on the page

`--addToQuickLaunch`
: Set, to add the page to the quick launch

--8<-- "docs/cmd/_global.md"

## Remarks

If you want to add the single-part app page to quick launch, use the addToQuickLaunch flag.

## Examples

Create a single-part app page in a site with url https://contoso.sharepoint.com, webpart data is stored in the `$webPartData` variable

```sh
m365 spo apppage add --title "Contoso" --webUrl "https://contoso.sharepoint.com" --webPartData $webPartData --addToQuickLaunch
```
