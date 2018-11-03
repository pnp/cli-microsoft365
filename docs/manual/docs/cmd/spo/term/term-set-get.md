# spo term set get

Gets information about the specified taxonomy term set

## Usage

```sh
spo term set get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id [id]`|ID of the term set to retrieve. Specify `name` or `id` but not both
`-n, --name [name]`|Name of the term set to retrieve. Specify `name` or `id` but not both
`--termGroupId [termGroupId]`|ID of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both
`--termGroupName [termGroupName]`|Name of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To get information about a taxonomy term set, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples

Get information about a taxonomy term set using its ID

```sh
spo term set get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb --termGroupName PnPTermSets
```

Get information about a taxonomy term set using its name

```sh
spo term set get --name PnPTermSets --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```