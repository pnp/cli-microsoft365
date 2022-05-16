# spo roledefinition get

Gets specified role definition from web by id

## Usage

```sh
m365 spo roledefinition get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve the role definition

`-i, --id <id>`
: The Id of the role definition to retrieve.

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve the role definition for site _https://contoso.sharepoint.com/sites/project-x_ with id _1_

```sh
m365 spo roledefinition get --webUrl https://contoso.sharepoint.com/sites/project-x --id 1
```
