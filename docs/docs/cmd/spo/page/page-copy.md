# spo page copy

Creates a copy of a modern page or template

## Usage

```sh
m365 spo page copy [options]
```

## Options

`--sourceName <sourceName>`
: The name of the source file

`--targetUrl <targetUrl>`
: The URL of the target file. You can specify page's name or relative- or absolute URL

`--overwrite`
: Overwrite the target page when it already exists

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located

--8<-- "docs/cmd/_global.md"

## Remarks

If another page exists at the specified target location, copying the page will fail, unless you include the `--overwrite` option.

## Examples

Create a copy of the `home.aspx` page

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "home.aspx" --targetUrl "home-copy.aspx"
```

Overwrite the target page if it already exists

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "home.aspx" --targetUrl "home-copy.aspx" --overwrite
```

Create a copy of a page template

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "templates/PageTemplate.aspx" --targetUrl "page.aspx"
```

Copy the page to another site

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "templates/PageTemplate.aspx" --targetUrl "https://contoso.sharepoint.com/sites/team-b/sitepages/page.aspx"
```
