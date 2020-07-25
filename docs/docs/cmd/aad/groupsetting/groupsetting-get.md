# aad groupsetting get

Gets information about the particular group setting

## Usage

```sh
m365 aad groupsetting get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The ID of the group setting to retrieve

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get information about the group setting with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad groupsetting get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```