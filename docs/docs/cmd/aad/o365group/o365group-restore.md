# aad o365group restore

Restores a deleted Microsoft 365 Group

## Usage

```sh
m365 aad o365group restore [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The ID of the Microsoft 365 Group to restore

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Restores the Microsoft 365 Group with id _28beab62-7540-4db1-a23f-29a6018a3848_

```sh
m365 aad o365group restore --id 28beab62-7540-4db1-a23f-29a6018a3848
```
