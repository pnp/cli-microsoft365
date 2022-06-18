# spo roledefinition get

Gets specified role definition from web by id

## Usage

```sh
m365 spo roledefinition get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve role definition

`-i, --id <id>`
: Role definition id

--8<-- "docs/cmd/_global.md"

## Examples

Return role definitions for site _https://contoso.sharepoint.com/sites/project-x_ with id _1_

```sh
m365 spo roledefinition get --webUrl https://contoso.sharepoint.com/sites/project-x --id 1
```
