# azmgmt flow environment get

Gets information about the specified Microsoft Flow environment

## Usage

```sh
azmgmt flow environment get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|The name of the environment to get information about
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Azure Management Service, using the [azmgmt connect](../connect.md) command.

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

To get information about the specified Microsoft Flow environment, you have to first connect to the Azure Management Service using the [azmgmt connect](../connect.md) command.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

## Examples

Get information about the Microsoft Flow environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_

```sh
azmgmt flow environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```