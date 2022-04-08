# spo serviceprincipal grant add

Grants the service principal permission to the specified API

## Usage

```sh
m365 spo serviceprincipal grant add [options]
```

## Alias

```sh
m365 spo sp grant add
```

## Options

`-r, --resource <resource>`
: The name of the resource for which permissions should be granted

`-s, --scope <scope>`
: The name of the permission that should be granted

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Grant the service principal permission to read email using the Microsoft Graph

```sh
m365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'Mail.Read'
```

Grant the service principal permission to a custom API

```sh
m365 spo serviceprincipal grant add --resource 'contoso-api' --scope 'user_impersonation'
```
