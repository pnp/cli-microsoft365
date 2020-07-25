# spo feature disable

Disables a feature for the specified site or web

## Usage

```sh
m365 spo feature disable [options]
```

## Options

Option|Description
------|-----------
`-h, --help`|output usage information
`-u, --url <url>`|The URL of the site or web for which to disable the feature
`-f, --featureId <id>`|The ID of the feature to disable
`-s, --scope [scope]`|Scope of the Feature to disable. Allowed values `Site,Web`. Default `Web`
`--force`|Specifies whether to continue if an error occurs when disabling the feature
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

If the specified url doesn't refer to an existing site collection, you will get a `"404 FILE NOT FOUND"` error.

## Examples

Disable site feature

```sh
m365 spo feature disable --url https://contoso.sharepoint.com/sites/sales --featureId 915c240e-a6cc-49b8-8b2c-0bff8b553ed3 --scope Site
```

Disable web feature (with force to ignore errors)

```sh
m365 spo feature disable --url https://contoso.sharepoint.com/sites/sales --featureId 00bfea71-5932-4f9c-ad71-1557e5751100 --scope Web --force
```
