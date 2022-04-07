# spo term group get

Gets information about the specified taxonomy term group

## Usage

```sh
m365 spo term group get [options]
```

## Options

`-i, --id [id]`
: ID of the term group to retrieve. Specify `name` or `id` but not both

`-n, --name [name]`
: Name of the term group to retrieve. Specify `name` or `id` but not both

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Get information about a taxonomy term group using its ID

```sh
m365 spo term group get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Get information about a taxonomy term group using its name

```sh
m365 spo term group get --name PnPTermSets
```
