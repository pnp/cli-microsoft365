# spo serviceprincipal set

Enable or disable the service principal

## Usage

```sh
spo serviceprincipal set [options]
```

## Alias

```sh
spo sp set
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-e, --enabled <enabled>`|Set to `true` to enable the service principal or to `false` to disable it. Valid values are `true|false`
`--confirm`|Don't prompt for confirming enabling/disabling the service principal
`-o, --output <output>`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To enable or disable the service principal, you have to first connect to a SharePoint tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

Using the `-e, --enabled` option you can specify whether the service principal should be enabled or disabled. Use `true` to enable the service principal and `false` to disable it.

## Examples

Enable the service principal. Will prompt for confirmation

```sh
spo serviceprincipal set --enabled true
```

Disable the service principal. Will prompt for confirmation

```sh
spo serviceprincipal set --enabled false
```

Enable the service principal without prompting for confirmation

```sh
spo serviceprincipal set --enabled true --confirm
```