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
`-n, --name [name]`|Display name for the Microsoft Teams team. Required, when `groupId` is not specified.
`-d, --description  [description]`|Description for the Microsoft Teams team. Required, when `groupId` is not specified.
`-i, --groupId [groupId]`|The ID of the Office 365 group to add a Microsoft Teams team to
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Add a new Microsoft Teams team by creating a group

```sh
teams team add --name 'Architecture' --description 'Architecture Discussion'
```

Add a new Microsoft Teams team to an existing Office 365 group

```sh
teams team add --groupId 6d551ed5-a606-4e7d-b5d7-36063ce562cc
```