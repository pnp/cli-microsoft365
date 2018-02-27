# spo web get

Retrieve information about the specified site

## Usage

```sh
spo web get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site for which to retrieve the information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To retrieve information about a site, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Retrieve information about the site _https://contoso.sharepoint.com/subsite_

```sh
spo web get --webUrl https://contoso.sharepoint.com/subsite
```