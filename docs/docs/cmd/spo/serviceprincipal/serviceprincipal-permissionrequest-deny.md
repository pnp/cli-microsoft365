# spo serviceprincipal permissionrequest deny

Denies the specified permission request

## Usage

```sh
m365 spo serviceprincipal permissionrequest deny [options]
```

## Alias

```sh
m365 spo sp permissionrequest deny
```

## Options

`-h, --help`
: output usage information

`-i, --requestId <requestId>`
: ID of the permission request to deny

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

## Remarks

The permission request you want to approve is denoted using its `ID`. You can retrieve it using the [spo serviceprincipal permissionrequest list](./serviceprincipal-permissionrequest-list.md) command.

## Examples

Deny permission request with id _4dc4c043-25ee-40f2-81d3-b3bf63da7538_

```sh
m365 spo serviceprincipal permissionrequest deny --requestId 4dc4c043-25ee-40f2-81d3-b3bf63da7538
```