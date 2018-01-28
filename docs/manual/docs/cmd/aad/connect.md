# aad connect

Connects to the Azure Active Directory Graph

## Usage

```sh
aad connect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

Using the `aad connect` command you can connect to the Azure Active Directory Graph to manage your AAD objects.

The `aad connect` command uses device code OAuth flow to connect to the AAD Graph.

When connecting to the AAD Graph, the `aad connect` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [aad disconnect](disconnect.md) command.

## Examples

Connect to the AAD Graph

```sh
aad connect
```

Connect to the AAD Graph in debug mode including detailed debug information in the console output

```sh
aad connect --debug
```