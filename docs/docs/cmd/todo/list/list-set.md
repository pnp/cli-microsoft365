# todo list set

Updates a Microsoft To Do task list

## Usage

```sh
m365 todo list set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id [id]`
: The ID of the list to update. Specify either id or name, not both

`-n, --name [name]`
: The display name of the list to update. Specify either id or name, not both

`--newName <newName>`
: The new name for the task list

`--query [query]`
: JMESPath query string. See <http://jmespath.org/> for more information and examples

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

Rename the list with ID _AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=_ to "My updated task list"

```sh
m365 todo list set --id "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=" --newName "My updated task list"
```

Rename the list with name _My Task list_ to "My updated task list"

```sh
m365 todo list set --name "My Task list" --newName "My updated task list"
```
