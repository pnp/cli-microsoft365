# spo site get

Gets information about the specific site collection

## Usage

```sh
spo site get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|URL of the site collection to retrieve information for
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To get information about a site collection, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

This command can retrieve information for both classic and modern sites.

## Examples

Return information about the _https://contoso.sharepoint.com/sites/project-x_ site collection.

```sh
spo site get -u https://contoso.sharepoint.com/sites/project-x
```
