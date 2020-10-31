# yammer message get

Returns a Yammer message

## Usage

```sh
m365 yammer message get [options]
```

## Options

`-h, --help`
: output usage information

`--id <id>`
: The id of the Yammer message

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
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Returns the Yammer message with the id 1239871123

```sh
m365 yammer message get --id 1239871123
```

Returns the Yammer message with the id 1239871123 in JSON format

```sh
m365 yammer message get --id 1239871123 --output json
```
