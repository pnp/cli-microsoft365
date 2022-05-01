# spo roledefinition list

Gets list of role definitions from web

## Usage

```sh
m365 spo roledefinition list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the role definitions to retrieve are located

--8<-- "docs/cmd/_global.md"

## Examples

Return list of role definitions located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo roledefinition list --webUrl https://contoso.sharepoint.com/sites/project-x
```
