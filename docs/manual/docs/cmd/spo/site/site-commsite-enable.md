# spo site commsite enable

Enables communication site features on the specified site

## Usage

```sh
spo site commsite enable [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|The URL of the site to enable communication site features on
`-i, --designPackageId [designPackageId]`|The ID of the site design to apply when enabling communication site features
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To enable communication site features on an existing site, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`. If you are logged in to a different site you will get an error.

## Examples

Enable communication site features on an existing site

```sh
spo site commsite enable --url https://contoso.sharepoint.com
```