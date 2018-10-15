# spo term add

Adds taxonomy term

## Usage

```sh
spo term add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the term to add
`--termSetId [termSetId]`|ID of the term set in which to create the term. Specify `termSetId` or `termSetName` but not both
`--termSetName [termSetName]`|Name of the term set in which to create the term. Specify `termSetId` or `termSetName` but not both
`--termGroupId [termGroupId]`|ID of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both
`--termGroupName [termGroupName]`|Name of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both
`-i, --id [id]`|ID of the term to add
`-d, --description [description]`|Description of the term to add
`--customProperties [customProperties]`|JSON string with key-value pairs representing custom properties to set on the term
`--localCustomProperties [localCustomProperties]`|JSON string with key-value pairs representing local custom properties to set on the term
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To add a taxonomy term, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

Add taxonomy term with the specified name to the term group and term set specified by their names

```sh
spo term add --name IT --termSetName Department --termGroupName People
```

Add taxonomy term with the specified name to the term group and term set specified by their IDs

```sh
spo term add --name IT --termSetId 8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f --termGroupId 5c928151-c140-4d48-aab9-54da901c7fef
```

Add taxonomy term with the specified name and ID

```sh
spo term add --name IT --id 5c928151-c140-4d48-aab9-54da901c7fef --termSetName Department --termGroupName People
```

Add taxonomy term with custom properties

```sh
spo term add --name IT --termSetName Department --termGroupName People --customProperties '{"Property": "Value"}'
```
