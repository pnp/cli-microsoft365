# todo list set

Updates a Microsoft To Do task list

## Usage

```sh
m365 todo list set [options]
```

## Options

`-i, --id [id]`
: The ID of the list to update. Specify either `id` or `name`, but not both.

`-n, --name [name]`
: The display name of the list to update. Specify either `id` or `name`, but not both.

`--newName <newName>`
: The new name for the task list.

--8<-- "docs/cmd/_global.md"

## Examples

Rename the list with a specific ID

```sh
m365 todo list set --id "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=" --newName "My updated task list"
```

Rename a list with a specific title

```sh
m365 todo list set --name "My Task list" --newName "My updated task list"
```

## Response

The command won't return a response on success.
