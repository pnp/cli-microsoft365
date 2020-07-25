# spo feature enable

Enables feature for the specified site or web

## Usage

```sh
m365 spo feature enable [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: The URL of the site or web for which to enable the feature

`-f, --featureId <id>`
: The ID of the feature to enable

`-s, --scope [scope]`
: Scope of the Feature to enable. Allowed values `Site,Web`. Default `Web`

`--force`
: Specifies whether to overwrite an existing feature with the same feature identifier. This parameter is ignored if there are no errors.

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified url doesn't refer to an existing site collection, you will get a `"404 FILE NOT FOUND"` error.

## Examples

Enable site feature

```sh
m365 spo feature enable --url https://contoso.sharepoint.com/sites/sales --featureId 915c240e-a6cc-49b8-8b2c-0bff8b553ed3 --scope Site
```

Enable web feature (with force to overwrite feature with same id)

```sh
m365 spo feature enable --url https://contoso.sharepoint.com/sites/sales --featureId 00bfea71-5932-4f9c-ad71-1557e5751100 --scope Web --force
```
