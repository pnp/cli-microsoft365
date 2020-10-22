# spo page get

Gets information about the specific modern page

## Usage

```sh
m365 spo page get [options]
```

## Options

`-n, --name <name>`
: Name of the page to retrieve

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified name doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

## Examples

Get information about the modern page with name _home.aspx_

```sh
m365 spo page get --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```