# spo web reindex

Requests reindexing the specified subsite

## Usage

```sh
spo web reindex [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the subsite to reindex
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To request reindexing a subsite, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If the subsite to be reindexed is a no-script site, the command will request reindexing all lists from the subsite that haven't been excluded from the search index.

## Examples

Request reindexing the subsite _https://contoso.sharepoint.com/subsite_

```sh
spo web reindex --webUrl https://contoso.sharepoint.com/subsite
```