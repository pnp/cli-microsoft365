# spo feature enable

Enable feature for the specified site or web

## Usage

```sh
spo feature enable [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|URL of the site (collection) to retrieve the activated Features from
`-f, --featureId <id>`|The ID of the feature to enable
`-s, --scope [scope]`|Scope of the Features to retrieve. Allowed values `Site|Web`. Default `Web`
`--force`|Specifies whether to overwrite an existing feature with the same feature identifier. This parameter is ignored if there are no errors.
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Enable feature on site

```sh
spo feature enable --url https://contoso.sharepoint.com/sites/sales --featureId 915c240e-a6cc-49b8-8b2c-0bff8b553ed3 --scope Site
```

Enable feature on web (with force to overwrite feature with same id)

```sh
spo feature enable --url https://contoso.sharepoint.com/sites/sales --featureId 00bfea71-5932-4f9c-ad71-1557e5751100 --scope Web --force
```