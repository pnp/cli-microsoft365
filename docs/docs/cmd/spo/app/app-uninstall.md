# spo app uninstall

Uninstalls an app from the site

## Usage

```sh
m365 spo app uninstall [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: ID of the app to uninstall

`-s, --siteUrl <siteUrl>`
: Absolute URL of the site to uninstall the app from

`--scope [scope]`
: Scope of the app catalog: `tenant,sitecollection`. Default `tenant`

`--confirm`
: Don't prompt for confirming uninstalling the app

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

Uninstall the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ from the _https://contoso.sharepoint.com_ site.

```sh
m365 spo app uninstall --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com
```

Uninstall the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ from the _https://contoso.sharepoint.com_ site without prompting for confirmation.

```sh
m365 spo app uninstall --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com
```

Uninstall the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ from the _https://contoso.sharepoint.com_ site where the app is deployed to the site collection app catalog of _https://contoso.sharepoint.com_.

```sh
m365 spo app uninstall --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com --scope sitecollection
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)