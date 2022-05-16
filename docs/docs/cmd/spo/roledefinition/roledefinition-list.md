# spo roledefinition list

Gets list of role definitions for the specified site

## Usage

```sh
m365 spo roledefinition list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve role definitions

--8<-- "docs/cmd/_global.md"

## Examples

Return list of role definitions for site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo roledefinition list --webUrl https://contoso.sharepoint.com/sites/project-x
```
