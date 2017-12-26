# spo serviceprincipal permissionrequest list

Lists pending permission requests

## Usage

```sh
spo serviceprincipal permissionrequest list [options]
```

## Alias

```sh
spo sp permissionrequest list
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output <output>`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To list pending permission requests, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples

List all pending permission requests

```sh
spo serviceprincipal permissionrequest list
```