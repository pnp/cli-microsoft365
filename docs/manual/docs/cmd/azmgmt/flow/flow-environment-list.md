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

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

List Microsoft Flow environments in the current tenant

```sh
azmgmt flow environment list
```