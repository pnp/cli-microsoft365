# aad o365group list

Lists Office 365 Groups in the current tenant

## Usage

```sh
aad o365group list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-d, --displayName [displayName]`|Retrieve only groups with displayName starting with the specified value
`-m, --mailNickname [displayName]`|Retrieve only groups with mailNickname starting with the specified value
`--includeSiteUrl`|Set to retrieve the site URL for each group
`--deleted`|Set to only retrieve deleted groups
`--orphaned`|Set to only retrieve groups without owners
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

Using the `--includeSiteUrl` option, you can retrieve the URL of the site associated with the particular Office 365 Group. If you however retrieve too many groups and will try to get their site URLs, you will most likely get an error as the command will get throttled, issuing too many requests, too frequently. If you get an error, consider narrowing down the result set using the `--displayName` and `--mailNickname` filters.

Retrieving the URL of the site associated with the particular Office 365 Group is not possible when retrieving deleted groups.

Using the `--orphaned` option, you can retrieve Office 365 Groups without owners.

## Examples

List all Office 365 Groups in the tenant

```sh
aad o365group list
```

List Office 365 Groups with display name starting with _Project_

```sh
aad o365group list --displayName Project
```

List Office 365 Groups mail nick name starting with _team_

```sh
aad o365group list --mailNickname team
```

List deleted Office 365 Groups with display name starting with _Project_

```sh
aad o365group list --displayName Project --deleted
```

List deleted Office 365 Groups mail nick name starting with _team_

```sh
aad o365group list --mailNickname team --deleted
```

List Office 365 Groups with display name starting with _Project_ including
the URL of the corresponding SharePoint site

```sh
aad o365group list --displayName Project --includeSiteUrl
```

List Office 365 Groups without owners

```sh
aad o365group list --orphaned
```
