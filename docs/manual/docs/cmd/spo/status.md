# spo status

Shows SharePoint Online site connection status

## Usage

```sh
spo status [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--verbose`|Runs command with verbose logging

## Remarks

If you are connected to a SharePoint Online, the spo status command
will show you information about the site to which you are connected, the currently stored
refresh and access token and the expiration date and time of the access token.

## Examples

```sh
spo status
```

shows the information about the current connection to SharePoint Online