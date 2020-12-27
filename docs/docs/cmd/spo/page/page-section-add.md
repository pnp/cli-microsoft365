# spo page section add

Adds section to modern page

## Usage

```sh
m365 spo page section add [options]
```

## Options

-n`, --name <name>`
: Name of the page to add section to

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located

`-t, --sectionTemplate <sectionTemplate>`
: Type of section to add. Allowed values `OneColumn,OneColumnFullWidth,TwoColumn,ThreeColumn,TwoColumnLeft,TwoColumnRight`

`--order [order]`
: Order of the section to add

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `name` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

Add section to the modern page named _home.aspx_

```sh
m365 spo page section add --name home.aspx --webUrl https://contoso.sharepoint.com/sites/newsletter  --sectionTemplate OneColumn --order 1
```