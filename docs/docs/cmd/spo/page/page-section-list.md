# spo page section list

List sections in the specific modern page

## Usage

```sh
m365 spo page section list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located

`-n, --name <name>`
: Name of the page to list sections of

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified name doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

List sections of a modern page named _home.aspx_

```sh
m365 spo page section list --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```