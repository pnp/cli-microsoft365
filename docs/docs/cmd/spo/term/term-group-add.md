# spo term group add

Adds taxonomy term group

## Usage

```sh
m365 spo term group add [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: Name of the term group to add

`-i, --id [id]`
: ID of the term group to add

`-d, --description [description]`
: Description of the term group to add

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

Add a new taxonomy term group with the specified name

```sh
m365 spo term group add --name PnPTermSets
```

Add a new taxonomy term group with the specified name and id

```sh
m365 spo term group add --name PnPTermSets --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Add a new taxonomy term group with the specified name and description

```sh
m365 spo term group add --name PnPTermSets --description 'Term sets for PnP'
```