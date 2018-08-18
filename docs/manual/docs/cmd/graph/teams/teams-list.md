# graph teams list

Lists Microsoft Teams teams in the current tenant

## Usage

```sh
graph teams list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-j, --joined`|Show only joined teams
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To list available Microsoft Teams teams, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

You can only see the details or archived status of the Microsoft Teams you are a member of.

## Examples

List all Microsoft Teams in the tenant

```sh
graph teams list
```

List all Microsoft Teams in the tenant you are a member of

```sh
graph teams list --joined
```