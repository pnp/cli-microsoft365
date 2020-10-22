# spo site inplacerecordsmanagement set

Activates or deactivates in-place records management for a site collection

## Usage

```sh
m365 spo site inplacerecordsmanagement set [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: The URL of the site on which to activate or deactivate in-place records management

`--enabled <enabled>`
: Set to `true` to activate in-place records management and to `false` to deactivate it

--8<-- "docs/cmd/_global.md"

## Examples

Activates in-place records management for site _https://contoso.sharepoint.com/sites/team-a_

```sh
m365 spo site inplacerecordsmanagement set --siteUrl https://contoso.sharepoint.com/sites/team-a --enabled true
```

Deactivates in-place records management for site _https://contoso.sharepoint.com/sites/team-a_

```sh
m365 spo site inplacerecordsmanagement set --siteUrl https://contoso.sharepoint.com/sites/team-a --enabled false
```