# cli config reset

Resets the specified CLI configuration option to its default value

## Usage

```sh
m365 cli config reset [options]
```

## Options

`-k, --key [key]`
: Config key to reset. If not specified, will reset all configuration settings to default

--8<-- "docs/cmd/_global.md"

--8<-- "docs/_clisettings.md"

## Examples

Reset CLI configuration option _showHelpOnFailure_ to its default value

```sh
m365 cli config reset --key showHelpOnFailure
```

Reset all configuration settings to default

```sh
m365 cli config reset
```

## More information

- [Configuring the CLI for Microsoft 365](../../../user-guide/configuring-cli.md)
