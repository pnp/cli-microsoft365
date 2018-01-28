# graph status

Shows Microsoft Graph connection status

## Usage

```sh
graph status [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

If you are connected to the Microsoft Graph, the `graph status` command will show you information about the currently stored refresh and access token and the expiration date and time of the access token when run in debug mode.

## Examples

Show the information about the current connection to the Microsoft Graph

```sh
graph status
```
