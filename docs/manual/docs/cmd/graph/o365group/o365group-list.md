# graph o365group list

Lists Office 365 Groups in the current tenant

## Usage

```sh
graph o365group list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-d, --displayName [displayName]`|Retrieve only groups with displayName starting with the specified value
`-m, --mailNickname [displayName]`|Retrieve only groups with mailNickname starting with the specified value
`--includeSiteUrl`|Set to retrieve the site URL for each group
`--deleted`|Set to only retrieve deleted groups
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To list available Office 365 Groups, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

Using the `--includeSiteUrl` option, you can retrieve the URL of the site associated with the particular Office 365 Group. If you however retrieve too many groups and will try to get their site URLs, you will most likely get an error as the command will get throttled, issuing too many requests, too frequently. If you get an error, consider narrowing down the result set using the `--displayName` and `--mailNickname` filters.

Retrieving the URL of the site associated with the particular Office 365 Group is not possible when retrieving deleted groups.

## Examples

List all Office 365 Groups in the tenant

```sh
graph o365group list
```

List Office 365 Groups with display name starting with _Project_

```sh
graph o365group list --displayName Project
```

List Office 365 Groups mail nick name starting with _team_

```sh
graph o365group list --mailNickname team
```

List deleted Office 365 Groups with display name starting with _Project_

```sh
graph o365group list --displayName Project --deleted
```

List deleted Office 365 Groups mail nick name starting with _team_

```sh
graph o365group list --mailNickname team --deleted
```

List Office 365 Groups with display name starting with _Project_ including
the URL of the corresponding SharePoint site

```sh
graph o365group list --displayName Project --includeSiteUrl
```