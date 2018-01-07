# spo serviceprincipal permissionrequest deny

Denies the specified permission request

## Usage

```sh
spo serviceprincipal permissionrequest deny [options]
```

## Alias

```sh
spo sp permissionrequest deny
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --requestId <requestId>`|ID of the permission request to deny
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To deny a permission request, you have to first connect to a tenant admin site using the
[spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

The permission request you want to approve is denoted using its `ID`. You can retrieve it using the [spo serviceprincipal permissionrequest list](./serviceprincipal-permissionrequest-list.md) command.

## Examples

Deny permission request with id _4dc4c043-25ee-40f2-81d3-b3bf63da7538_

```sh
spo serviceprincipal permissionrequest deny --requestId 4dc4c043-25ee-40f2-81d3-b3bf63da7538
```