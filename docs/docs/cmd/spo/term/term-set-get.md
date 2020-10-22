# spo term set get

Gets information about the specified taxonomy term set

## Usage

```sh
m365 spo term set get [options]
```

## Options

`-i, --id [id]`
: ID of the term set to retrieve. Specify `name` or `id` but not both

`-n, --name [name]`
: Name of the term set to retrieve. Specify `name` or `id` but not both

`--termGroupId [termGroupId]`
: ID of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both

`--termGroupName [termGroupName]`
: Name of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Get information about a taxonomy term set using its ID

```sh
m365 spo term set get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb --termGroupName PnPTermSets
```

Get information about a taxonomy term set using its name

```sh
m365 spo term set get --name PnPTermSets --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```