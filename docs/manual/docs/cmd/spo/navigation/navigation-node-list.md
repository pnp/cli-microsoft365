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
`-l, --location <location>`|Navigation type to retrieve. Available options: `QuickLaunch`, `TopNavigationBar`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To retrieve navigation nodes for a site, you have to first log in to a SharePoint Online site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Retrieve nodes from the top navigation

```sh
spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar
```

Retrieve nodes from the quick launch

```sh
spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch
```