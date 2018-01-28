# graph connect

Connects to the Microsoft Graph

## Usage

```sh
graph connect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

Using the `graph connect` command you can connect to the Microsoft Graph.

The `graph connect` command uses device code OAuth flow to connect to the Microsoft Graph.

When connecting to the Microsoft Graph, the `graph connect` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [graph disconnect](disconnect.md) command.

## Examples

Connect to the Microsoft Graph

```sh
graph connect
```

Connect to the Microsoft Graph in debug mode including detailed debug information in the console output

```sh
graph connect --debug
```