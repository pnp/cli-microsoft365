# spo orgnewssite remove

Removes a site from the list of organizational news sites

## Usage

```sh
m365 spo orgnewssite remove [options]
```

## Options

`-u, --url <url>`
: Absolute URL of the site to remove

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Remove a site from the list of organizational news

```sh
m365 spo orgnewssite remove --url https://contoso.sharepoint.com/sites/site1
```

Remove a site from the list of organizational news sites, without prompting for confirmation

```sh
m365 spo orgnewssite remove --url https://contoso.sharepoint.com/sites/site1 --confirm
```
