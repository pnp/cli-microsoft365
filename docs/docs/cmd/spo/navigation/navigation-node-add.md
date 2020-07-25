# spo navigation node add

Adds a navigation node to the specified site navigation

## Usage

```sh
m365 spo navigation node add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site to which navigation should be modified

`-l, --location <location>`
: Navigation type where the node should be added. Available options: `QuickLaunch`, `TopNavigationBar`

`-t, --title <title>`
: Navigation node title

`--url <url>`
: Navigation node URL

`--parentNodeId [parentNodeId]`
: ID of the node below which the node should be added

`--isExternal`
: Set, if the navigation node points to an external URL

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Add a navigation node pointing to a SharePoint page to the top navigation

```sh
m365 spo navigation node add --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar --title About --url /sites/team-s/sitepages/about.aspx
```

Add a navigation node pointing to an external page to the quick launch

```sh
m365 spo navigation node add --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch --title "About us" --url https://contoso.com/about-us --isExternal
```

Add a navigation node below an existing node

```sh
m365 spo navigation node add --webUrl https://contoso.sharepoint.com/sites/team-a --parentNodeId 2010 --title About --url /sites/team-s/sitepages/about.aspx
```