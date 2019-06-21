# graph planner task list

Lists Planner tasks of the current logged in user

## Usage

```sh
graph planner task list
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging.
`--debug`|Runs command with debug logging.

!!! important
    before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To list planner tasks of a current logged in user, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

If you are not assigned with any task it will return empty results.

## Examples

List all the tasks of current logged in user

```sh
graph planner task list
```

## More information

- Microsoft Graph Get Tasks of User: 
[https://docs.microsoft.com/en-us/graph/api/planneruser-list-tasks?view=graph-rest-1.0&tabs=cs](https://docs.microsoft.com/en-us/graph/api/planneruser-list-tasks?view=graph-rest-1.0&tabs=cs)