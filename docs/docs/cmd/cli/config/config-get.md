# cli config get

Gets value of a CLI for Microsoft 365 configuration option

## Usage

```sh
m365 cli config get [options]
```

## Options

`-k, --key <key>`
: Config key to get the value of

--8<-- "docs/cmd/_global.md"

--8<-- "docs/_clisettings.md"

## Remarks

If the specified setting has not been configured, CLI will return no output.

## Examples

Get the output configured for CLI for Microsoft 365

```sh
m365 cli config get --key output
```

## More information

- [Configuring the CLI for Microsoft 365](../../../user-guide/configuring-cli.md)
