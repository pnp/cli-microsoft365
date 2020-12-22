# spo web get

Retrieve information about the specified site

## Usage

```sh
m365 spo web get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve the information

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve information about the site _https://contoso.sharepoint.com/subsite_

```sh
m365 spo web get --webUrl https://contoso.sharepoint.com/subsite
```