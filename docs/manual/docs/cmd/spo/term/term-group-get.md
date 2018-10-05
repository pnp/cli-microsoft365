# spo term group get

Gets information about the specified taxonomy term group

## Usage

```sh
spo term group get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id [id]`|ID of the term group to retrieve. Specify `name` or `id` but not both
`-n, --name [name]`|Name of the term group to retrieve. Specify `name` or `id` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To get information about a taxonomy term group, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

Get information about a taxonomy term group using its ID

```sh
spo term group get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Get information about a taxonomy term group using its name

```sh
spo term group get --name PnPTermSets
```