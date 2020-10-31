# spo navigation node remove

Removes the specified navigation node

## Usage

```sh
m365 spo navigation node remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site to which navigation should be modified

`-l, --location <location>`
: Navigation type where the node should be added. Available options: `QuickLaunch,TopNavigationBar`

`-i, --id <id>`
: ID of the node to remove

`--confirm`
: Don't prompt for confirming removing the node

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Remove a node from the top navigation. Will prompt for confirmation

```sh
m365 spo navigation node remove --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar --id 2003
```

Remove a node from the quick launch without prompting for confirmation

```sh
m365 spo navigation node remove --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch --id 2003 --confirm
```