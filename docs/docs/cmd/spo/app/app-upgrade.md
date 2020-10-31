# spo app upgrade

Upgrades app in the specified site

## Usage

```sh
m365 spo app upgrade [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: ID of the app to upgrade

`-s, --siteUrl <siteUrl>`
: Absolute URL of the site to upgrade the app in

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

If the app with the specified ID doesn't exist in the app catalog, the command will fail with an error.

## Examples

Upgrade the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ in the _https://contoso.sharepoint.com_ site.

```sh
m365 spo app upgrade --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com
```

Upgrade the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ in the _https://contoso.sharepoint.com_ site from site collection app catalog.

```sh
m365 spo app upgrade --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com --scope sitecollection
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)