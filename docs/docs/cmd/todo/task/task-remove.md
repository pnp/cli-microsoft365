# todo task remove

Removes the specified Microsoft To Do task

## Usage

```sh
m365 todo task remove [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The id of the task to remove

`--listName [listName]`
: The name of the task list in which the task exists. Specify either `listId` or `listName`, not both

`--listId [listId]`
: The id of the task list in which the task exists. Specify either `listId` or `listName`, not both

`--confirm`
: Don't prompt for confirmation

`--query [query]`
: JMESPath query string. See http://jmespath.org/ for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Removes Microsoft To Do task with the specified id in a list with the specified name

```sh
m365 todo task remove --id "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=" --listName "Tasks"
```

Removes Microsoft To Do task with the specified id in a list with the specified id

```sh
m365 todo task remove --id "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=" --listId "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA="
```
