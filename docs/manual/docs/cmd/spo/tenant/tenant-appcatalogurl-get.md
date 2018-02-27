# spo tenant appcatalogurl get

Gets the URL of the tenant app catalog

## Usage

```sh
spo tenant appcatalogurl get [options]
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

## Examples

Get the URL of the tenant app catalog

```sh
spo tenant appcatalogurl get
```