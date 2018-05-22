# graph groupsetting list

Lists Azure AD group settings

## Usage

```sh
graph groupsetting list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To list group settings, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

## Examples

List all group settings in the tenant

```sh
graph groupsetting list
```