# graph groupsettingtemplate list

Lists Azure AD group settings templates

## Usage

```sh
graph groupsettingtemplate list [options]
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

To list group setting templates, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

## Examples

List all group setting templates in the tenant

```sh
graph groupsettingtemplate list
```