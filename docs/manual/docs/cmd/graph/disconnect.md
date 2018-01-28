# graph disconnect

Disconnects from the Microsoft Graph

## Usage

```sh
graph disconnect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

The `graph disconnect` command disconnects from the Microsoft Graph and removes any access and refresh tokens from memory

## Examples

Disconnect from Microsoft Graph

```sh
graph disconnect
```

Disconnect from Microsoft Graph in debug mode including detailed debug information in the console output

```sh
graph disconnect --debug
```