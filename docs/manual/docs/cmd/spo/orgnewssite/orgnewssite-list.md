# spo orgnewssite list

Lists all organizational news sites

## Usage

```sh
spo orgnewssite list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To list all sites identified as organizational news sites, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`

## Examples

List all organizational news sites

```sh
spo orgnewssite list
```