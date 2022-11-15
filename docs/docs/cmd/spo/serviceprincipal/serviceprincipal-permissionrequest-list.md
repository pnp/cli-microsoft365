# spo serviceprincipal permissionrequest list

Lists pending permission requests

## Usage

```sh
m365 spo serviceprincipal permissionrequest list [options]
```

## Alias

```sh
m365 spo sp permissionrequest list
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    The admin role that's required to approve/deny permissions depends on the API. To approve permissions to any of the third-party APIs registered in the tenant, the application administrator role is sufficient. To approve permissions for Microsoft Graph or any other Microsoft API, the Global Administrator role is required.

## Examples

List all pending permission requests

```sh
m365 spo serviceprincipal permissionrequest list
```
