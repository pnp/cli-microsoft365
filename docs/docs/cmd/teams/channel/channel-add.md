# teams channel add

Adds a channel to the specified Microsoft Teams team

## Usage

```sh
m365 teams channel add [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the team to add the channel to

`-n, --name <name>`
: The name of the channel to add

`-d, --description [description]`
: The description of the channel to add

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

You can only add a channel to the Microsoft Teams team you are a member of.

## Examples

Add channel to the specified Microsoft Teams team

```sh
m365 teams channel add --teamId 6703ac8a-c49b-4fd4-8223-28f0ac3a6402 --name office365cli --description development
```