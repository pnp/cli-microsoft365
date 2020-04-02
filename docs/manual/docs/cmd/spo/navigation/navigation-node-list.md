# spo navigation node list

Lists nodes from the specified site navigation

## Usage

```sh
spo navigation node list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|Absolute URL of the site for which to retrieve navigation
`-l, --location <location>`|Navigation type to retrieve. Available options: `QuickLaunch`&#x7c;`TopNavigationBar`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. <code>json&124;text</code>. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Retrieve nodes from the top navigation

```sh
spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar
```

Retrieve nodes from the quick launch

```sh
spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch
```