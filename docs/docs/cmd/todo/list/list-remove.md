# todo list remove

Removes a Microsoft To Do task list

## Usage

```sh
m365 todo list remove [options]
```

## Options

`-n, --name [name]`
: The name of the task list to remove. Specify either `id` or `name`, but not both.

`-i, --id [id]`
: The ID of the task list to remove. Specify either `id` or `name`, but not both.

`--confirm`
: Don't prompt for confirming removing the task list.

--8<-- "docs/cmd/_global.md"

## Examples

Remove a task list with specific name

```sh
m365 todo list remove --name "My task list"
```

Remove a task list with the ID without confirmation prompt

```sh
m365 todo list remove --id "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=" --confirm
```

## Response

The command won't return a response on success.
