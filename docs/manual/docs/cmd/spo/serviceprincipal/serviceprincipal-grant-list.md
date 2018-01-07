# spo serviceprincipal grant list

Lists permissions granted to the service principal

## Usage

```sh
spo serviceprincipal grant list [options]
```

## Alias

```sh
spo sp grant list
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To list permission granted to the service principal, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples

List all permissions granted to the service principal

```sh
spo serviceprincipal grant list
```