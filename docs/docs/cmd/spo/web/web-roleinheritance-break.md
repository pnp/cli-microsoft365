# spo web roleinheritance break

Break role inheritance of subsite.

## Usage

```sh
m365 spo web roleinheritance break [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site

--8<-- "docs/cmd/_global.md"

## Examples

Break role inheritance of subsite _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo web roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x
```