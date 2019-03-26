# spo web list

Lists subsites of the specified site

## Usage

```sh
spo web list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the parent site for which to retrieve the list of subsites
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Return all subsites from site _https://contoso.sharepoint.com/_

```sh
spo web list -u https://contoso.sharepoint.com
```