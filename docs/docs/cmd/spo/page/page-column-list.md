# spo page column list

Lists columns in the specific section of a modern page

## Usage

```sh
m365 spo page column list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located

`-n, --name <name>`
: Name of the page to list columns of

`-s, --section <sectionId>`
: ID of the section for which to list columns

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified name doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

List columns in the first section of a modern page with name _home.aspx_

```sh
m365 spo page column list --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx --section 1
```
