# spo term set list

Lists taxonomy term sets from the given term group

## Usage

```sh
spo term set list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--termGroupId [termGroupId]`|ID of the term group from which to retrieve term sets. Specify `termGroupName` or `termGroupId` but not both
`--termGroupName [termGroupName]`|Name of the term group from which to retrieve term sets. Specify `termGroupName` or `termGroupId` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To list taxonomy term sets, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

List taxonomy term sets from the term group with the given name

```sh
spo term set list --termGroupName PnPTermSets
```

List taxonomy term sets from the term group with the given ID

```sh
spo term set list --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```