# spo term group list

Lists taxonomy term groups

## Usage

```sh
spo term group list [options]
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

To list taxonomy term groups, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples

List taxonomy term groups

```sh
spo term group list
```