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
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To list group setting templates, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

List all group setting templates in the tenant

```sh
graph groupsettingtemplate list
```