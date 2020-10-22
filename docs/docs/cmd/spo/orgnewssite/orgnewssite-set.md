# spo orgnewssite set

Marks site as an organizational news site

## Usage

```sh
m365 spo orgnewssite set [options]
```

## Options

`-u, --url <url>`
: The URL of the site to mark as an organizational news site

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-u, --url` option you can specify which site to add to the list of organizational news sites.

## Examples

Set a site as an organizational news site

```sh
m365 spo orgnewssite set --url https://contoso.sharepoint.com/sites/site1
```
