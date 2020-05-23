# todo list add

Adds a new Microsoft To Do task list

## Usage

```sh
todo list add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name`|The name of the task list to add
`--query [query]`|JMESPath query string. See http://jmespath.org/ for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Add a list called "My task list"
      
```sh
todo list add --name "My task list"
```