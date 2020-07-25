# flow environment get

Gets information about the specified Microsoft Flow environment

## Usage

```sh
m365 flow environment get [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: The name of the environment to get information about

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

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

## Examples

Get information about the Microsoft Flow environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_

```sh
m365 flow environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```