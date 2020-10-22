# spo web list

Lists subsites of the specified site

## Usage

```sh
m365 spo web list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the parent site for which to retrieve the list of subsites

--8<-- "docs/cmd/_global.md"

## Examples

Return all subsites from site _https://contoso.sharepoint.com/_

```sh
m365 spo web list -u https://contoso.sharepoint.com
```