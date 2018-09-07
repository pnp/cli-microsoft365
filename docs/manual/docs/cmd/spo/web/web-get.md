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
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To retrieve information about a site, you have to first log in to a SharePoint Online site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Retrieve information about the site _https://contoso.sharepoint.com/subsite_

```sh
spo web get --webUrl https://contoso.sharepoint.com/subsite
```