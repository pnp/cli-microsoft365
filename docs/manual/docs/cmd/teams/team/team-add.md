# teams team add

Adds a new Microsoft Teams team

## Usage

```sh
teams team add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Display name for the Microsoft Teams team
`-d, --description <description>`|Description for the Microsoft Teams team
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If you want to add a Team to an existing Office 365 Group use the [aad o365group teamify](../../aad/o365group/o365group-teamify.md) command instead.

## Examples

Add a new Microsoft Teams team

```sh
teams team add --name 'Architecture' --description 'Architecture Discussion'
```