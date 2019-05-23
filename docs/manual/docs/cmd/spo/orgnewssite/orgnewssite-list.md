# spo orgnewssite list

Lists URLs of all organizational news sites for the tenant

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
If you are logged in to a different site and try to manage tenant properties,
you will get an error.

## Examples

List all organizational news sites

```sh
spo orgnewssite list
```