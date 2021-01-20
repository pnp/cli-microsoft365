# spo page copy

Creates a copy of a modern page or template.

## Usage

```sh
m365 spo page copy [options]
```

## Options

`--sourceName <sourceName>`
: The name of the source file

`--sourceName <sourceName>`
: The name of the target file to create

`--overwrite`
: Overwrite the target page when it already exists

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located

--8<-- "docs/cmd/_global.md"

## Examples

Create a copy of the `home.aspx` page.

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "home.aspx" --targetName "home-copy.aspx"
```

Overwrite the page copy if it already exists.

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "home.aspx" --targetName "home-copy.aspx" --overwrite
```

Create a copy of a page template.

```sh
m365 spo page copy --webUrl https://contoso.sharepoint.com/sites/team-a --sourceName "templates/PageTemplate.aspx" --targetName "page.aspx"
```