# azmgmt connect

Connects to the Azure Management Service

## Usage

```sh
azmgmt connect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

Using the `azmgmt connect` command you can connect to the Azure Management Service to manage your Azure objects.

The `azmgmt connect` command uses device code OAuth flow to connect to the Azure Management Service.

When connecting to the Azure Management Service, the `azmgmt connect` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [azmgmt disconnect](disconnect.md) command.

## Examples

Connect to the Azure Management Service

```sh
azmgmt connect
```

Connect to the Azure Management Service in debug mode including detailed debug information in the console output

```sh
azmgmt connect --debug
```