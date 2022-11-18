# spo serviceprincipal permissionrequest deny

Denies the specified permission request

## Usage

```sh
m365 spo serviceprincipal permissionrequest deny [options]
```

## Alias

```sh
m365 spo sp permissionrequest deny
```

## Options

`-i, --id <id>`
: ID of the permission request to deny

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    The admin role that's required to deny permissions depends on the API. To approve permissions to any of the third-party APIs registered in the tenant, the application administrator role is sufficient. To approve permissions for Microsoft Graph or any other Microsoft API, the Global Administrator role is required.

The permission request you want to approve is denoted using its `ID`. You can retrieve it using the [spo serviceprincipal permissionrequest list](./serviceprincipal-permissionrequest-list.md) command.

## Examples

Deny permission request with id _4dc4c043-25ee-40f2-81d3-b3bf63da7538_

```sh
m365 spo serviceprincipal permissionrequest deny --id 4dc4c043-25ee-40f2-81d3-b3bf63da7538
```
