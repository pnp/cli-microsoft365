# spo term list

Lists taxonomy terms from the given term set

## Usage

```sh
spo term list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--termGroupId [termGroupId]`|ID of the term group where the term set is located. Specify `termGroupId` or `termGroupName` but not both
`--termGroupName [termGroupName]`|Name of the term group where the term set is located. Specify `termGroupId` or `termGroupName` but not both
`--termSetId [termSetId]`|ID of the term set for which to retrieve terms. Specify `termSetId` or `termSetName` but not both
`--termSetName [termSetName]`|Name of the term set for which to retrieve terms. Specify `termSetId` or `termSetName` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To list taxonomy terms, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples

List taxonomy terms from the term group and term set with the given name

```sh
spo term list --termGroupName PnPTermSets --termSetName PnP-Organizations
```

List taxonomy terms from the term group and term set with the given ID

```sh
spo term list --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb --termSetId 0e8f395e-ff58-4d45-9ff7-e331ab728bec
```