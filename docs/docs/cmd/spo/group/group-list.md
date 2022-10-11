# spo group list

Lists all the groups within specific web

## Usage

```sh
m365 spo group list [options]
```

## Options

`-u, --webUrl <webUrl>`
: Url of the web to list the group within

`--associatedGroupsOnly`
: Get only the associated visitor, member and owner groups of the site.

--8<-- "docs/cmd/_global.md"

## Examples

Lists all the groups within a specific web

```sh
m365 spo group list --webUrl "https://contoso.sharepoint.com/sites/contoso"
```

Lists the associated groups within a specific web

```sh
m365 spo group list --webUrl "https://contoso.sharepoint.com/sites/contoso" --associatedGroupsOnly
```
