# todo list list

Returns a list of Microsoft To Do task lists

## Usage

```sh
todo list list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See http://jmespath.org/ for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Get the list of Microsoft To Do task lists

```sh
todo list list
```
