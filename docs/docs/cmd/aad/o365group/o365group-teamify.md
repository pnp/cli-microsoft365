# aad o365group teamify

Creates a new Microsoft Teams team under existing Microsoft 365 group

## Usage

```sh
m365 aad o365group teamify [options]
```

## Options

`-h, --help`
: output usage information

`-i, --groupId <groupId>`
: The ID of the Microsoft 365 Group to connect to Microsoft Teams

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Creates a new Microsoft Teams team under existing Microsoft 365 group

```sh
m365 aad o365group teamify --groupId e3f60f99-0bad-481f-9e9f-ff0f572fbd03
```