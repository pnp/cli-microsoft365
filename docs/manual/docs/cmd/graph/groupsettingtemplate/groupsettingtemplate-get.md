# graph groupsettingtemplate get

Gets information about the specified Azure AD group settings template

## Usage

```sh
graph groupsettingtemplate get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id [id]`|The ID of the settings template to retrieve. Specify the `id` or `displayName` but not both
`-n, --displayName [displayName]`|The display name of the settings template to retrieve. Specify the `id` or `displayName` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To get information about a group setting template, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

## Examples

Get information about the group setting template with id _62375ab9-6b52-47ed-826b-58e47e0e304b_

```sh
graph groupsettingtemplate get --id 62375ab9-6b52-47ed-826b-58e47e0e304b
```

Get information about the group setting template with display name _Group.Unified_

```sh
graph groupsettingtemplate get --displayName Group.Unified
```