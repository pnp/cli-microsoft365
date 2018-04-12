# azmgmt disconnect

Disconnects from the Azure Management Service

## Usage

```sh
azmgmt disconnect [options]
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

The `azmgmt disconnect` command disconnects from the Azure Management Service and removes any access and refresh tokens from memory.

## Examples

Disconnect from the Azure Management Service

```sh
azmgmt disconnect
```

Disconnect from the Azure Management Service in debug mode including detailed debug information in the console output

```sh
azmgmt disconnect --debug
```