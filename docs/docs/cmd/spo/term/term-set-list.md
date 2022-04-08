# spo term set list

Lists taxonomy term sets from the given term group

## Usage

```sh
m365 spo term set list [options]
```

## Options

`--termGroupId [termGroupId]`
: ID of the term group from which to retrieve term sets. Specify `termGroupName` or `termGroupId` but not both

`--termGroupName [termGroupName]`
: Name of the term group from which to retrieve term sets. Specify `termGroupName` or `termGroupId` but not both

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

List taxonomy term sets from the term group with the given name

```sh
m365 spo term set list --termGroupName PnPTermSets
```

List taxonomy term sets from the term group with the given ID

```sh
m365 spo term set list --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```
