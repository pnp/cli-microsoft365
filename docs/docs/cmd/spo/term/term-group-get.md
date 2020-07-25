# spo term group get

Gets information about the specified taxonomy term group

## Usage

```sh
m365 spo term group get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id [id]`
: ID of the term group to retrieve. Specify `name` or `id` but not both

`-n, --name [name]`
: Name of the term group to retrieve. Specify `name` or `id` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Get information about a taxonomy term group using its ID

```sh
m365 spo term group get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Get information about a taxonomy term group using its name

```sh
m365 spo term group get --name PnPTermSets
```