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
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To list permission granted to the service principal, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

List all permissions granted to the service principal

```sh
spo serviceprincipal grant list
```