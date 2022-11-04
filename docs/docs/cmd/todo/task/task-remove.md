# todo task remove

Removes the specified Microsoft To Do task

## Usage

```sh
m365 todo task remove [options]
```

## Options

`-i, --id <id>`
: The id of the task to remove.

`--listName [listName]`
: The name of the task list in which the task exists. Specify either `listId` or `listName`, but not both.

`--listId [listId]`
: The id of the task list in which the task exists. Specify either `listId` or `listName`, but not both.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Removes Microsoft To Do task with the specified id in a list with the specified name

```sh
m365 todo task remove --id "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=" --listName "Tasks"
```

Removes Microsoft To Do task with the specified id in a list with the specified id

```sh
m365 todo task remove --id "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=" --listId "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA="
```

## Response

The command won't return a response on success.

