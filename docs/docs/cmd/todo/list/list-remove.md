# todo list remove

Removes a Microsoft To Do task list

## Usage

```sh
m365 todo list remove [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name [name]`
: The name of the task list to remove. Specify either `id` or `name` but not both

`-i, --id [id]`
: The ID of the task list to remove. Specify either `id` or `name` but not both

`--confirm`
: Don't prompt for confirming removing the task list

`--query [query]`
: JMESPath query string. See http://jmespath.org/ for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Remove a task list with the name _My task list_

```sh
m365 todo list remove --name "My task list"
```

Remove a task list with the ID _AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=_

```sh
m365 todo list remove --id "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA="
```
