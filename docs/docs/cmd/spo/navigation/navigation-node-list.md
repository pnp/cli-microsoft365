# spo navigation node list

Lists nodes from the specified site navigation

## Usage

```sh
m365 spo navigation node list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site for which to retrieve navigation

`-l, --location <location>`
: Navigation type to retrieve. Available options: `QuickLaunch,TopNavigationBar`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Retrieve nodes from the top navigation

```sh
m365 spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar
```

Retrieve nodes from the quick launch

```sh
m365 spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch
```