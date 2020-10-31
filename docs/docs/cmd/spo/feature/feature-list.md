# spo feature list

Lists Features activated in the specified site or site collection

## Usage

```sh
m365 spo feature list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: URL of the site (collection) to retrieve the activated Features from

`-s, --scope [scope]`
: Scope of the Features to retrieve. Allowed values `Site,Web`. Default `Web`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Return details about Features activated in the specified site collection

```sh
m365 spo feature list --url https://contoso.sharepoint.com/sites/test --scope Site
```

Return details about Features activated in the specified site

```sh
m365 spo feature list --url https://contoso.sharepoint.com/sites/test --scope Web
```