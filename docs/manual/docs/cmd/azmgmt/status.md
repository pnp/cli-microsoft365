# azmgmt status

Shows Azure Management Service connection status

## Usage

```sh
azmgmt status [options]
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

If you are connected to the Azure Management Service, the `azmgmt status` command will show you information about the currently stored refresh and access token and the expiration date and time of the access token when run in debug mode. If you are connected using a user name and password, it will also show you the name of the user used to authenticate.

## Examples

Show the information about the current connection to the Azure Management Service

```sh
azmgmt status
```