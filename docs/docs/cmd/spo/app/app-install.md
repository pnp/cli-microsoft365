# spo app install

Installs an app from the specified app catalog in the site

## Usage

```sh
m365 spo app install [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: ID of the app to install

`-s, --siteUrl <siteUrl>`
: Absolute URL of the site to install the app in

`--scope [scope]`
: Scope of the app catalog: `tenant,sitecollection`. Default `tenant`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the app with the specified ID doesn't exist in the app catalog, the command will fail with an error. Before you can install app in a site, you have to add it to the tenant or site collection app catalog first using the [spo app add](./app-add.md) command.

## Examples

Install the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ in the _https://contoso.sharepoint.com_ site.

```sh
m365 spo app install --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com
```

Install the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ in the _https://contoso.sharepoint.com_ site from site collection app catalog.

```sh
m365 spo app install --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com --scope sitecollection
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)