# graph planner task list

Lists Planner tasks for the currently logged in user

## Usage

```sh
graph planner task list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging.
`--debug`|Runs command with debug logging.

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To list Planner tasks for the currently logged in user, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

List tasks for the currently logged in user

```sh
graph planner task list
```
