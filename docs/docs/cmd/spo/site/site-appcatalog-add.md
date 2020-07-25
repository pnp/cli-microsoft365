# spo site appcatalog add

Creates a site collection app catalog in the specified site

## Usage

```sh
m365 spo site appcatalog add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: URL of the site collection where the app catalog should be added

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Add a site collection app catalog to the specified site

```sh
m365 spo site appcatalog add --url https://contoso.sharepoint/sites/site
```

## More information

- Use the site collection app catalog: [https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog](https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog)