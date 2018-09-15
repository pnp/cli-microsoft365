# azmgmt logout

Log out from the Azure Management Service

## Usage

```sh
azmgmt logout [options]
```

## Alias

```sh
azmgmt disconnect
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
    The 'azmgmt disconnect' command is deprecated. Please use 'azmgmt logout' instead.

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

The `azmgmt logout` command logs out from the Azure Management Service and removes any access and refresh tokens from memory.

## Examples

Log out from the Azure Management Service

```sh
azmgmt logout
```

Log out from the Azure Management Service in debug mode including detailed debug information in the console output

```sh
azmgmt logout --debug
```