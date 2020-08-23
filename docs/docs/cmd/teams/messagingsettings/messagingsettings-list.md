# teams messagingsettings list

Lists messaging settings for a Microsoft Teams team

## Usage

```sh
teams messagingsettings list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId`|The ID of the team for which to get the messaging settings
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get messaging settings for a Microsoft Teams team

```sh
teams messagingsettings list --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
```