# spo status

Shows SharePoint Online site login status

## Usage

```sh
spo status [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

If you are logged in to SharePoint Online, the spo status command will show you information about the site to which you are logged in, the currently stored refresh and access token and the expiration date and time of the access token.

## Examples

Show the information about the current login to SharePoint Online

```sh
spo status
```
