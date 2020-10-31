# spo sitedesign run status get

Gets information about the site scripts executed for the specified site design

## Usage

```sh
m365 spo sitedesign run status get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site for which to get the information

`-i, --runId <runId>`
: ID of the site design applied to the site as retrieved using `spo sitedesign run list`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

For text output mode, displays the name of the action, site script and the outcome of the action. For JSON output mode, displays all available information.

## Examples

List information about site scripts executed for the specified site design

```sh
m365 spo sitedesign run status get --webUrl https://contoso.sharepoint.com/sites/team-a --runId b4411557-308b-4545-a3c4-55297d5cd8c8
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)