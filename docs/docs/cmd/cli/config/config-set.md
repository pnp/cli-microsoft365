# cli config set

Manage global configuration settings about the CLI for Microsoft 365.

## Usage

```sh
m365 cli config set [options]
```

## Options

`-k, --key <key>`
: Config key to . Allowed values: `showHelpOnFailure`

`-v, --value <value>`
: Config value to set

--8<-- "docs/cmd/_global.md"

## Remarks

Using the `cli config set` command you can set CLI for Microsoft 365 settings.

## Examples

Set configuration to always display help on command execution failure

```sh
m365 cli config set --key showHelpOnFailure --value true
```
