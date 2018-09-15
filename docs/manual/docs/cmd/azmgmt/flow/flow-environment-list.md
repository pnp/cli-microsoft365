# azmgmt flow environment list

Lists Microsoft Flow environments in the current tenant

## Usage

```sh
azmgmt flow environment list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Azure Management Service, using the [azmgmt login](../login.md) command.

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

To get information about Microsoft Flow environments, you have to first log in to the Azure Management Service using the [azmgmt login](../login.md) command.

## Examples

List Microsoft Flow environments in the current tenant

```sh
azmgmt flow environment list
```