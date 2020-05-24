# todo list remove

Removes a Microsoft To Do task list

## Usage

```sh
todo list remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name [name]`|The name of the task list to remove. Specify either id or name but not both
`-i, --id [id]`|The ID of the list to remove. Specify either id or name but not both
`--confirm`|Don't prompt for confirming removing the task list
`--query [query]`|JMESPath query string. See http://jmespath.org/ for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Remove the list "My task list"

```sh
todo list remove --name "My task list"
```

Remove the list with Id "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA="

```sh
todo list remove --id "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA="
```