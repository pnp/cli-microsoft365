# pa app list

Lists all Power Apps apps

## Usage

```sh
pa app list [options]
```

## Options

Option|Description
------|-----------
`-h, --help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reaches general availability.

## Examples

List all apps

```sh
pa app list
```
