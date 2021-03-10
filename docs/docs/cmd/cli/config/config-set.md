# cli config set

Sets CLI for Microsoft 365 configuration options

## Usage

```sh
m365 cli config set [options]
```

## Options

`-k, --key <key>`
: Config key to specify. Allowed values: `showHelpOnFailure`

`-v, --value <value>`
: Config value to set

--8<-- "docs/cmd/_global.md"

## Examples

Configure CLI to automatically display help when executing a command failed

```sh
m365 cli config set --key showHelpOnFailure --value true
```

## More information

- [Configuring the CLI for Microsoft 365](../../../user-guide/configuring-cli.md)