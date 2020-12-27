# spo navigation node list

Lists nodes from the specified site navigation

## Usage

```sh
m365 spo navigation node list [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site for which to retrieve navigation

`-l, --location <location>`
: Navigation type to retrieve. Available options: `QuickLaunch,TopNavigationBar`

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve nodes from the top navigation

```sh
m365 spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar
```

Retrieve nodes from the quick launch

```sh
m365 spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch
```