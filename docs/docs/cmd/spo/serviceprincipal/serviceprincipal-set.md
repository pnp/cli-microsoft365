# spo serviceprincipal set

Enable or disable the service principal

## Usage

```sh
m365 spo serviceprincipal set [options]
```

## Alias

```sh
m365 spo sp set
```

## Options

`-h, --help`
: output usage information

`-e, --enabled <enabled>`
: Set to `true` to enable the service principal or to `false` to disable it. Valid values are `true,false`

`--confirm`
: Don't prompt for confirming enabling/disabling the service principal

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

Using the `-e, --enabled` option you can specify whether the service principal should be enabled or disabled. Use `true` to enable the service principal and `false` to disable it.

## Examples

Enable the service principal. Will prompt for confirmation

```sh
m365 spo serviceprincipal set --enabled true
```

Disable the service principal. Will prompt for confirmation

```sh
m365 spo serviceprincipal set --enabled false
```

Enable the service principal without prompting for confirmation

```sh
m365 spo serviceprincipal set --enabled true --confirm
```