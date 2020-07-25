# pa connector list

Lists custom connectors in the given environment

## Usage

```sh
m365 pa connector list [options]
```

## Alias

```sh
m365 flow connector list
```

## Options

`-h, --help`
: output usage information

`-e, --environment <environment>`
: The name of the environment for which to retrieve custom connectors

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

List all custom connectors in the given environment

```sh
m365 pa connector list --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```
