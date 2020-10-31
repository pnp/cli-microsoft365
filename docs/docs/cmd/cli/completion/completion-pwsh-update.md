# cli completion pwsh update

Updates command completion for PowerShell

## Usage

```sh
m365 cli completion pwsh update [options]
```

## Options

`-h, --help`
: output usage information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

This commands updates the list of commands and their options used by command completion in PowerShell. You should run this command each time after upgrading the CLI for Microsoft 365.

## Examples

Update list of commands for PowerShell command completion

```powershell
cli completion pwsh update
```

## More information

- Command completion: [https://pnp.github.io/cli-microsoft365/concepts/completion/](https://pnp.github.io/cli-microsoft365/concepts/completion/)
