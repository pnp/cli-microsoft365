# spo serviceprincipal permissionrequest list

Lists pending permission requests

## Usage

```sh
spo serviceprincipal permissionrequest list [options]
```

## Alias

```sh
spo sp permissionrequest list
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

List all pending permission requests

```sh
spo serviceprincipal permissionrequest list
```