# spo group list

Lists all the groups within specific web

## Usage

```sh
m365 spo group list [options]
```

## Options

`-u, --webUrl <webUrl>`
: Url of the web to list the group within

--8<-- "docs/cmd/_global.md"

## Examples

Lists all the groups within specific web _https://contoso.sharepoint.com/sites/contoso_

```sh
m365 spo group list --webUrl "https://contoso.sharepoint.com/sites/contoso"
```
