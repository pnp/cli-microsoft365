# spo site appcatalog add

Creates a site collection app catalog in the specified site

## Usage

```sh
m365 spo site appcatalog add [options]
```

## Options

`-u, --url <url>`
: URL of the site collection where the app catalog should be added

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Add a site collection app catalog to the specified site

```sh
m365 spo site appcatalog add --url https://contoso.sharepoint/sites/site
```

## More information

- Use the site collection app catalog: [https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog](https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog)