# spo web installedlanguage list

Lists all installed languages on site

## Usage

```sh
m365 spo web installedlanguage list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve the list of installed languages

--8<-- "docs/cmd/_global.md"

## Examples

Return all installed languages from site _https://contoso.sharepoint.com/_

```sh
m365 spo web installedlanguage list --webUrl https://contoso.sharepoint.com
```
