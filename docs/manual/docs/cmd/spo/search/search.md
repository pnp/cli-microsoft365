# spo search

Execute a search query

## Usage

```sh
spo search [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-q, --query <query>`|Query to execute
`-p, --selectProperties`|Comma separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested.
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To execute a search query, you have to first log in to a SharePoint Online site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Execute search query to retrieve all Document Sets (ContentTypeId = '0x0120D520')

```sh
spo search --query 'ContentTypeId:0x0120D520'
```

Retrieve all documents. For each document, retrieve the Path, Author and FileType.

```sh
spo search --query 'IsDocument:1' --selectProperties 'Path,Author,FileType' --allResults
```