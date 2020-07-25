# spo group list

Lists all the groups within specific web

## Usage

```sh
m365 spo group list [options]
```

## Options

Option|Description
------|-----------
`-h, --help`|output usage information
`-u, --webUrl <webUrl>`|Url of the web to list the group within
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output

## Examples

Lists all the groups within specific web _https://contoso.sharepoint.com/sites/contoso_

```sh
m365 spo group list --webUrl "https://contoso.sharepoint.com/sites/contoso"
```