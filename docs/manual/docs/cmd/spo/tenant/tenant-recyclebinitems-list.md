# spo tenant recyclebinitems list

Returns all modern and classic site collections in the tenant scoped recycle bin

## Usage

```sh
spo tenant recyclebinitems list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Returns all modern and classic site collections in the tenant scoped recycle bin

```sh
spo tenant recyclebinitems list
```