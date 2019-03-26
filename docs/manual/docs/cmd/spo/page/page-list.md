# spo page list

Lists all modern pages in the given site

## Usage

```sh
spo page list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site from which to retrieve available pages
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

List all modern pages in the specific site

```sh
spo page list --webUrl https://contoso.sharepoint.com/sites/team-a
```