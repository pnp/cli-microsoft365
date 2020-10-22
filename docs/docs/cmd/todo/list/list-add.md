# todo list add

Adds a new Microsoft To Do task list

## Usage

```sh
m365 todo list add [options]
```

## Options

`-n, --name <name>`
: The name of the task list to add

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Add a task list with the name _My task list_

```sh
m365 todo list add --name "My task list"
```
